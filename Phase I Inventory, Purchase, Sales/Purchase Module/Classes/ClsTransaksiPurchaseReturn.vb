Option Explicit On
Option Strict Off

Public Class ClsTransaksiPurchaseReturn
    Private mPrimary As Double = -1
    Private mNoTransaksi As String = ""
    Private mNoUrut As Double = -1
    Private mPrimaryStockOutPurchaseReturn As Double = -1
    Private mPrimaryGRN As Double = -1
    Private mTanggalTransaksi As Date = Now
    Private mSubTotal As Double = 0
    Private mTaxExclude As Boolean = True
    Private mTax As Double = 0
    Private mTaxAmount As Double = 0
    Private mTotal As Double = 0
    Private mReasonOfReturn As String = ""
    Private mRemarks As String = ""
    Private mSupplierDONumber As String = ""
    Private mPrimarySupplier As Double = -1
    Private mNoNPWPSupplier As String = ""
    Private mVATDipungut As Boolean = True
    Private mNoNotaRetur As String = ""
    Private mInvoiceCancellation As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Dim mTipe_Return As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Revisi_AOP
    Dim mPrimaryCurrency As Double = -1
    Dim mDutystamp As Double = 0
    Dim mPaymentTerm As Double = 0
    Dim mTaxExchangeRate As Double = 1

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian
    Dim ObjProduct As New Core_Module.clsMasterProduct
    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code mulai()
    '====================================================================================================================================================================
    Private mPrimaryAccount As Double = -1
    Private mPrimaryCurrTax As Double = -1

    Public Property InvoiceCancellation() As Integer
        Get
            Return mInvoiceCancellation
        End Get
        Set(ByVal value As Integer)
            mInvoiceCancellation = value
        End Set
    End Property


    Public Function GetPrimaryAccountAP(ByVal vPrimaryPR As Decimal) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryAccount(vPrimaryPR).HasValue AndAlso DTA.ScalarQuery_GetPrimaryAccount(vPrimaryPR).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryAccount(vPrimaryPR).Value
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


    Public Function GetPrimaryCurrTax(ByVal vPrimaryPR As Decimal) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter


        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryPR).HasValue AndAlso DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryPR).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimaryCurrTax(vPrimaryPR).Value
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

    Public Property PrimaryAccount() As Double
        Get
            Return Me.mPrimaryAccount
        End Get
        Set(ByVal value As Double)
            mPrimaryAccount = value
        End Set
    End Property

    Public Property PrimaryCurrTax() As Double
        Get
            Return Me.mPrimaryCurrTax
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrTax = value
        End Set
    End Property

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
  ByVal vDgvItem As DataGridView, ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn3DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn3Row
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem

        Try
            'fil basis data
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.Primary = .Primary
                    Me.NoTransaksi = .No_Transaksi
                    Me.PrimaryStockOutPurchaseReturn = .Primary_Stock_Out_Purchase_Return
                    Me.PrimaryGRN = .Primary_GRN
                    Me.TanggalTransaksi = .Tanggal_Transaksi
                    Me.SubTotal = .Sub_Total
                    Me.TaxExclude = .Tax_Exclude
                    Me.Tax = .Tax
                    Me.TaxAmount = .Tax_Amount
                    Me.Total = .Total
                    Me.ReasonOfReturn = .Reason_Of_Return
                    Me.Remarks = .Remarks
                    Me.SupplierDONumber = .Supplier_DO_Number

                    Me.mDutystamp = .Dutystamp
                    Me.mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mPaymentTerm = .Payment_Term
                    Me.mTipe_Return = .Tipe_PR
                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_ap_account
                    Me.mNoNotaRetur = .No_Nota_Retur_Pajak
                    Me.mNoNPWPSupplier = .No_NPWP
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDipungut = True
                    Else
                        Me.mVATDipungut = False
                    End If
                    Me.InvoiceCancellation = .Invoice_Cancellation

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV_Multi(vDgvItem, Me.mPrimary, vFrm)
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
    ByVal vDgvItem As DataGridView, ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, _
    ByRef vTBL As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable, _
    ByRef vTblCustomInOUt As CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn2TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn2DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn2Row
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturnQuick_AddByCustomProduct
        Try
            'fil basis data
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.Primary = .Primary
                    Me.NoTransaksi = .No_Transaksi
                    Me.PrimaryStockOutPurchaseReturn = .Primary_Stock_Out_Purchase_Return
                    Me.PrimaryGRN = .Primary_GRN
                    Me.TanggalTransaksi = .Tanggal_Transaksi
                    Me.SubTotal = .Sub_Total
                    Me.TaxExclude = .Tax_Exclude
                    Me.Tax = .Tax
                    Me.TaxAmount = .Tax_Amount
                    Me.Total = .Total
                    Me.ReasonOfReturn = .Reason_Of_Return
                    Me.Remarks = .Remarks
                    Me.SupplierDONumber = .Supplier_DO_Number

                    Me.mPrimarySupplier = .primary_supplier
                    Me.mDutystamp = .Dutystamp
                    Me.mTaxExchangeRate = .Tax_Exchange_Rate
                    Me.mPaymentTerm = .Payment_Term
                    Me.mTipe_Return = .Tipe_PR
                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mPrimaryCurrTax = .Primary_Curr_Tax
                    Me.mPrimaryAccount = .primary_ap_account

                    Me.mNoNotaRetur = .No_Nota_Retur_Pajak
                    Me.mNoNPWPSupplier = .No_NPWP
                    If .VAT_Dipungut = 1 Then
                        Me.mVATDipungut = True
                    Else
                        Me.mVATDipungut = False
                    End If

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV_Multi(vDgvItem, Me.mPrimary, vFrm)
                ObjCustom.IsiTblDummyFromDB(Me.mPrimary, vTBL, vTblCustomInOUt)
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

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vSetPOB As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim DTA3 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem
        Dim ObjGRN As New ClsTransaksiGRN
        Dim VAT_Int As Integer = 0

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            ''@rosadinnn
            ''@MUaT
            If Me.ApakahDataBenar(True, vdgvItem, vFrm.ColPrimaryCodeNo.Index, vFrm.ColPrimaryPO.Index, _
            vFrm.ColQtyMulti.Index, ObjGRN.GetPrimaryStockIn(Me.mPrimaryGRN), Me.mPrimaryGRN, Me.mPrimary, vCekDataBenar) = True Then
                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'insert query
                If Me.mVATDipungut = True Then
                    VAT_Int = 1
                Else
                    VAT_Int = 0
                End If

                If DTA3.InsertQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mPrimaryStockOutPurchaseReturn, _
                   Me.mPrimaryGRN, Me.mTanggalTransaksi, Me.mSubTotal, Me.mTaxExclude, _
                   Me.mTax, Me.mTaxAmount, Me.mTotal, Me.mReasonOfReturn, Me.mRemarks, CreatedBy_, _
                   CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mSupplierDONumber, Me.mTipe_Return, Me.mPrimaryCurrency, Me.mTaxExchangeRate, _
                   Me.mDutystamp, Me.mPaymentTerm, Me.mPrimarySupplier, Me.mNoNPWPSupplier, VAT_Int, Me.mNoNotaRetur, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_getPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm, vSetPOB)
                    If vFromCutOff = False Then
                        Me.CekNomor(DTA)
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
                Return Me.SimpanBaru1(vdgvItem, vUserCode, vFrm, vCekDataBenar, vSetPOB, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA3.Dispose()
            ObjItem = Nothing
            ObjGRN = Nothing
        End Try
    End Function


    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, _
    ByVal vTblCustom As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable, _
    Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal NoStokcOutPR As String = "", Optional ByVal vSetPOB As Boolean = True, _
    Optional ByVal vKoreksiProdQty As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim DTA3 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem
        Dim ObjGRN As New ClsTransaksiGRN
        Dim VAT_Int As Integer = 0

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            If Me.ApakahDataBenar(True, vdgvItem, vFrm, ObjGRN.GetPrimaryStockIn(Me.mPrimaryGRN), Me.mPrimaryGRN, Me.mPrimary, vCekDataBenar) = True Then
                If Me.mVATDipungut = True Then
                    VAT_Int = 1
                Else
                    VAT_Int = 0
                End If

                'insert query
                If DTA3.InsertQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mPrimaryStockOutPurchaseReturn, _
                   Me.mPrimaryGRN, Me.mTanggalTransaksi, Me.mSubTotal, Me.mTaxExclude, _
                   Me.mTax, Me.mTaxAmount, Me.mTotal, Me.mReasonOfReturn, Me.mRemarks, CreatedBy_, _
                   CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mSupplierDONumber, Me.mTipe_Return, Me.mPrimaryCurrency, Me.mTaxExchangeRate, _
                   Me.mDutystamp, Me.mPaymentTerm, Me.mPrimarySupplier, Me.mNoNPWPSupplier, VAT_Int, Me.mNoNotaRetur, _
                   Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_getPrimary(Me.mNoTransaksi).Value
                    'set sebagai quick purchase return
                    DTA.UpdateQuery_SetQuickPurchase(Me.mPrimary)

                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm, vTblCustom, vSetPOB)
                    If vFromCutOff = False Then
                        Me.CekNomor(DTA)
                    End If

                    'simpan data stock out purchase return langsung
                    Me.GenerateStockOutPurchaseReturn_QuickPR_Multi(vdgvItem, vFrm, NoStokcOutPR, vKoreksiProdQty, vSetPOB, vCekDataBenar, vFromCutOff)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru1(vdgvItem, vUserCode, vFrm, vTblCustom, vCekDataBenar, NoStokcOutPR, vSetPOB, vKoreksiProdQty, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjGRN = Nothing
        End Try
    End Function


    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim DTA3 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem
        Dim UseTax As Int16 = 0
        Dim ObjGRN As New ClsTransaksiGRN
        Dim VAT_Int As Integer = 0
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturn

        Try
            If Me.ApakahDataBenar(False, dgvItem, vFrm.ColPrimaryCodeNo.Index, vFrm.ColPrimaryPO.Index, _
            vFrm.ColQtyMulti.Index, ObjGRN.GetPrimaryStockIn(Me.mPrimaryGRN), Me.mPrimaryGRN, Me.mPrimary) = True Then
                If Me.TaxExclude = True Then
                    UseTax = 1
                Else
                    UseTax = 0
                End If
                If Me.mVATDipungut = True Then
                    VAT_Int = 1
                Else
                    VAT_Int = 0
                End If

                If DTA3.UpdateQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mPrimaryStockOutPurchaseReturn, _
                Me.mPrimaryGRN, Me.mTanggalTransaksi, Me.mSubTotal, Me.mTaxExclude, _
                Me.mTax, Me.mTaxAmount, Me.mTotal, Me.mReasonOfReturn, Me.mRemarks, vUserCode, Now, Now, Me.mSupplierDONumber, _
                Me.mTipe_Return, Me.mPrimaryCurrency, Me.mTaxExchangeRate, Me.mDutystamp, Me.mPaymentTerm, Me.mPrimarySupplier, _
                Me.mNoNPWPSupplier, VAT_Int, Me.mNoNotaRetur, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm)


                    ObjCustom.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)

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
                Return Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA3.Dispose()
            ObjItem = Nothing
            ObjGRN = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
 ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, ByVal vTblCustom As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim DTA3 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn3TableAdapter
        Dim ObjItem As New ClsTransaksiPurchaseReturnItem
        Dim UseTax As Int16 = 0
        Dim ObjGRN As New ClsTransaksiGRN
        Dim NoTransaksiStockOutLama As String = ""
        Dim VAT_Int As Integer = 0
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable 'kemungkinan dikirim dari form quick nya
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturn

        Try
            If Me.ApakahDataBenar(False, dgvItem, vFrm, ObjGRN.GetPrimaryStockIn(Me.mPrimaryGRN), Me.mPrimaryGRN, Me.mPrimary) = True Then
                'hapus stock out purchase return lama
                Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
                ObjStockInOut.CreatedBy = ObjStockInOut.GetCreatedBy(Me.mPrimaryStockOutPurchaseReturn)
                ObjStockInOut.Primary = Me.mPrimaryStockOutPurchaseReturn
                ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                If Format(Me.GetTanggalTransaksi(Me.mPrimary), "dd MMM yyyy") <> Format(Me.mTanggalTransaksi, "dd MMM yyyy") Then
                    NoTransaksiStockOutLama = ""
                Else
                    NoTransaksiStockOutLama = ObjStockInOut.GetNoTransaksi(Me.mPrimaryStockOutPurchaseReturn)
                End If
                If ObjStockInOut.Hapus1(Me.mPrimaryStockOutPurchaseReturn, Me.mCreatedBy, TblItemCustom) = False Then
                    ObjStockInOut = Nothing
                    Return False
                End If
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                ObjStockInOut = Nothing

                If Me.TaxExclude = True Then
                    UseTax = 1
                Else
                    UseTax = 0
                End If
                If Me.mVATDipungut = True Then
                    VAT_Int = 1
                Else
                    VAT_Int = 0
                End If

                If DTA3.UpdateQuery(Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 5)), Me.mPrimaryStockOutPurchaseReturn, _
                Me.mPrimaryGRN, Me.mTanggalTransaksi, Me.mSubTotal, Me.mTaxExclude, _
                Me.mTax, Me.mTaxAmount, Me.mTotal, Me.mReasonOfReturn, Me.mRemarks, vUserCode, Now, Now, Me.mSupplierDONumber, _
                Me.mTipe_Return, Me.mPrimaryCurrency, Me.mTaxExchangeRate, Me.mDutystamp, Me.mPaymentTerm, Me.mPrimarySupplier, _
                Me.mNoNPWPSupplier, VAT_Int, Me.mNoNotaRetur, Me.mPrimaryAccount, Me.mPrimaryCurrTax, Me.mInvoiceCancellation, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'set sebagai quick purchase return
                    DTA.UpdateQuery_SetQuickPurchase(Me.mPrimary)
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 5)
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode

                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary, Me)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm, vTblCustom)

                    'simpan stock out purchase return baru
                    Me.GenerateStockOutPurchaseReturn_QuickPR_Multi(dgvItem, vFrm, NoTransaksiStockOutLama)

                    ObjCustom.SimpanPerubahanCustom(Me.mPrimary, Me.mNoTransaksi)

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
                Return Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm, vTblCustom)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjGRN = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    '====================================================================================================================================================================
    '====================================================================================================================================================================
    'Erna's Code selesai()
    '====================================================================================================================================================================

    Public Property PrimarySupplier() As Double
        Get
            Return Me.mPrimarySupplier
        End Get
        Set(ByVal value As Double)
            Me.mPrimarySupplier = value
        End Set
    End Property

    Public Property TaxExchangeRate() As Double
        Get
            Return Me.mTaxExchangeRate
        End Get
        Set(ByVal value As Double)
            mTaxExchangeRate = value
        End Set
    End Property

    Public Property NoNPWPSupplier() As String
        Get
            Return Me.mNoNPWPSupplier
        End Get
        Set(ByVal value As String)
            mNoNPWPSupplier = value
        End Set
    End Property

    Public Property NoNotaReturPajak() As String
        Get
            Return Me.mNoNotaRetur
        End Get
        Set(ByVal value As String)
            mNoNotaRetur = value
        End Set
    End Property

    Public Property VATDipungut() As Boolean
        Get
            Return Me.mVATDipungut
        End Get
        Set(ByVal value As Boolean)
            mVATDipungut = value
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

    Public Property Dutystamp() As Double
        Get
            Return Me.mDutystamp
        End Get
        Set(ByVal value As Double)
            mDutystamp = value
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

    Public Property TipeReturn() As EnumTipe_PurchaseReturn
        Get
            Return Me.mTipe_Return
        End Get
        Set(ByVal value As EnumTipe_PurchaseReturn)
            mTipe_Return = value
        End Set
    End Property

    Public Enum EnumTipe_PurchaseReturn As Integer
        Revisi_AOP
        Nota_Retur
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property NoTransaksi() As String
        Get
            Return (mNoTransaksi)
        End Get
        Set(ByVal value As String)
            mNoTransaksi = value
        End Set
    End Property

    Public Property NoUrut() As Double
        Get
            Return mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
        End Set
    End Property

    Public Property PrimaryStockOutPurchaseReturn() As Double
        Get
            Return mPrimaryStockOutPurchaseReturn
        End Get
        Set(ByVal value As Double)
            mPrimaryStockOutPurchaseReturn = value
        End Set
    End Property

    Public Property PrimaryGRN() As Double
        Get
            Return mPrimaryGRN
        End Get
        Set(ByVal value As Double)
            mPrimaryGRN = value
        End Set
    End Property
    Public Property TanggalTransaksi() As Date
        Get
            Return mTanggalTransaksi
        End Get
        Set(ByVal value As Date)
            mTanggalTransaksi = value
        End Set
    End Property
    Public Property SubTotal() As Double
        Get
            Return mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
        End Set
    End Property
    Public Property TaxExclude() As Boolean
        Get
            Return mTaxExclude
        End Get
        Set(ByVal value As Boolean)
            mTaxExclude = value
        End Set
    End Property
    Public Property Tax() As Double
        Get
            Return mTax
        End Get
        Set(ByVal value As Double)
            mTax = value
        End Set
    End Property
    Public Property TaxAmount() As Double
        Get
            Return mTaxAmount
        End Get
        Set(ByVal value As Double)
            mTaxAmount = value
        End Set
    End Property
    Public Property Total() As Double
        Get
            Return mTotal
        End Get
        Set(ByVal value As Double)
            mTotal = value
        End Set
    End Property
    Public Property ReasonOfReturn() As String
        Get
            Return mReasonOfReturn
        End Get
        Set(ByVal value As String)
            mReasonOfReturn = value
        End Set
    End Property
    Public Property Remarks() As String
        Get
            Return mRemarks
        End Get
        Set(ByVal value As String)
            mRemarks = value
        End Set
    End Property
    Public Property SupplierDONumber() As String
        Get
            Return mSupplierDONumber
        End Get
        Set(ByVal value As String)
            mSupplierDONumber = value
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

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable4TableAdapter
        Try
            If DTA.ScalarQuery_getPrimaryCurrency(vPrimary).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_getPrimaryCurrency(vPrimary).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimaryPR As Double) As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If DTA.ScalarQuery_getTimeStamp(vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_getTimeStamp(vPrimaryPR).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetInfo(ByVal vPrimaryPR As Double, ByRef vSupplierDoNO As String, ByRef vNoStockInOut As String, ByRef vPrimaryStockInOut As String) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.DataTable4DataTable
        Dim Row As DseTransaksiPurchaseReturn.DataTable4Row
        Try
            vSupplierDoNO = ""
            vNoStockInOut = ""
            vPrimaryStockInOut = ""
            DTA.FillBy_PrimaryPR(TBL, vPrimaryPR)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                vSupplierDoNO = Row.Supplier_DO_Number
                vNoStockInOut = Row.No_Stock_In_Out
                vPrimaryStockInOut = Row.Primary_Stock_Out_Purchase_Return
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function
    Public Function GetInfo2(ByVal vNoPR As String, ByRef vSupplierDoNO As String, ByRef vNoStockInOut As String, ByRef vPrimaryStockInOut As String) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable4TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.DataTable4DataTable
        Dim Row As DseTransaksiPurchaseReturn.DataTable4Row
        Try
            vSupplierDoNO = ""
            vNoStockInOut = ""
            vPrimaryStockInOut = ""
            DTA.FillBy_NoPR(TBL, vNoPR)
            If TBL.Rows.Count > 0 Then
                Row = TBL.Rows(0)
                vSupplierDoNO = Row.Supplier_DO_Number
                vNoStockInOut = Row.No_Stock_In_Out
                vPrimaryStockInOut = Row.Primary_Stock_Out_Purchase_Return
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Di pakai sama AOP untuk load data purchase return yang akan dimasukan ke dalam AOP.
    ''' Purchase return yang boleh diload ke dalam AOP adalah purchase return yang:
    ''' 1. GRN yang dirujuk oleh PR tsb belum pernah diload ke dalam AOP
    ''' atau
    ''' 2. GRN yang dirujuk oleh PR tsb tidak pernah diload dalam AOPR
    ''' atau
    ''' 3. GRN yang dirujuk oleh PR tsb pernah diload ke dalam AOP yang primarynya diberikan
    ''' 4. Tipe PR tersebut adalah purchase return dengan nota return.
    ''' </summary>
    ''' <param name="ListGRN"></param>
    ''' <param name="TBL"></param>
    ''' <remarks></remarks>
    Public Sub IsiTBLPRByPrimaryGRN(ByVal ListGRN As List(Of Double), ByVal TBL As _
    DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1DataTable, ByVal vPrimaryAOP As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_view1TableAdapter
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_view1Row
        Dim counter As Double = 0
        Dim TipePR As EnumTipe_PurchaseReturn = EnumTipe_PurchaseReturn.Revisi_AOP

        Dim PrimaryGRN As Double = 0
        Try
            DTA.ClearBeforeFill = False
            TBL.Rows.Clear()
            'yang diload adalah semua purchase return yang tidak punya AOPR
            For Each PrimaryGRN In ListGRN
                DTA.FillBy_4AOP(TBL, PrimaryGRN, TipePR)
            Next
            'cek berkaitan dengan AOP yang merujuk ke PR tsb
            While counter < TBL.Rows.Count
                row = TBL.Rows(counter)
                If row.Primary_AOP <> -1 Then
                    If row.Primary_AOP <> vPrimaryAOP Then
                        'purchase return tersebut sudah dirujuk oleh AOP lain, 
                        'tidak boleh dimasukkan ke dalam daftar PR
                        TBL.Rows.RemoveAt(counter)
                    Else
                        'cek untuk item berikutnya
                        counter += 1
                    End If
                Else
                    'cek untuk item berikutnya
                    counter += 1
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiTBLPurchaseReturnBelumAdaAOPR(ByVal vPrimarySupplier As Double, _
           ByVal vTBL As DseTransaksiPurchaseReturn.DataTable8DataTable) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable8TableAdapter
        Try
            DTA.Fill(vTBL, vPrimarySupplier)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_getUrutanMax(TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_getUrutanMax(TanggalStart, TanggalEnd).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'pastikan no urut terdiri atas 5 huruf
            While NoUrut.Length < 5
                NoUrut = "0" + NoUrut
            End While

            'bulan 
            Bulan = CStr(vTanggalTransaksi.Month)
            While Bulan.Length < 2
                Bulan = "0" + Bulan
            End While
            Return NoUrut + "/PRTN/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If DTA.ScalarQuery_getJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_getJml(vNoTransaksi).Value > 0 Then
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

  
    Public Sub IsiTBL_PO_OB(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, ByVal vPrimaryStockInPurchase As Double, _
    ByVal vPrimaryGRN As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.PR_OBTableAdapter
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim TBL1 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter
        Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2Row
        Dim Filter As String = ""
        Dim Dummy As String = ""

        Try
            'isikan tabel stock in purchase
            DTA.Fill(vTBL, vPrimaryStockInPurchase)

            'kurangkan qty yang ada dengan qty yang sudah dimasukkan sebagai purchase return, sehingga yang didapatkan hanya order balancenya
            'saja
            DTA1.Fill(TBL1, vPrimaryGRN)
            Dummy = vTBL.DefaultView.RowFilter
            For Each row1 In TBL1.Rows
                Filter = "[Primary Product] = " + row1.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row1.Primary_PO.ToString
                vTBL.DefaultView.RowFilter = Filter
                While vTBL.DefaultView.Count > 0 And row1.Total_Qty > 0
                    row = vTBL.DefaultView(0).Row
                    If row.Qty > row1.Total_Qty Then
                        row.Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    Else
                        row1.Total_Qty -= row.Qty
                        vTBL.Rows.Remove(row)
                    End If
                End While
            Next

            vTBL.DefaultView.RowFilter = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Public Sub IsikanTBL_Product(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, ByVal vPrimaryStockInPurchase As Double, _
    ByVal vPrimaryGRN As Double, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.PR_OBTableAdapter
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim TBL1 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter
        Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2Row
        Dim Filter As String = ""
        Dim Dummy As String = ""

        Try
            'isikan tabel stock in purchase
            DTA.FillBy_PrimaryPO(vTBL, vPrimaryStockInPurchase, vPrimaryPO)

            'kurangkan qty yang ada dengan qty yang sudah dimasukkan sebagai purchase return, sehingga yang didapatkan hanya order balancenya
            'saja
            DTA1.FillBy_PrimaryPO_PrimaryGRN(TBL1, vPrimaryGRN, vPrimaryPO)
            Dummy = vTBL.DefaultView.RowFilter
            For Each row1 In TBL1.Rows
                Filter = "[Primary Product] = " + row1.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row1.Primary_PO.ToString
                vTBL.DefaultView.RowFilter = Filter
                While vTBL.DefaultView.Count > 0 And row1.Total_Qty > 0
                    row = vTBL.DefaultView(0).Row
                    If row.Qty > row1.Total_Qty Then
                        row.Qty -= row1.Total_Qty
                        row1.Total_Qty = 0
                    Else
                        row1.Total_Qty -= row.Qty
                        vTBL.Rows.Remove(row)
                    End If
                End While
            Next

            vTBL.DefaultView.RowFilter = Dummy
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLPR_ByPrimaryTransaksi(ByVal vTBL As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable, _
    ByVal vPrimaryPR As Double, ByVal vPrimaryGRN As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter

        Try
            DTA.FillBy_PrimaryPR(vTBL, vPrimaryGRN, vPrimaryPR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLPR_ByPrimaryTransaksi_PrimaryPO(ByVal vTBL As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable, _
  ByVal vPrimaryPR As Double, ByVal vPrimaryGRN As Double, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter

        Try
            DTA.FillBy_PrimaryPO_PrimaryPR(vTBL, vPrimaryGRN, vPrimaryPR, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function IsiDGV(ByVal vCriteria As String, _
       ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPurchaseReturn.DataTable3DataTable, _
       ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, ByVal IsQuick As Integer, Optional ByVal vStartDate As Date = (#1/1/1950#), Optional ByVal vEndDate As Date = (#1/1/2100#)) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TglDariSampai(TBL, vStartDate, vEndDate, IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, vCriteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SupplierNameLIKE(TBL, vCriteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SupplierCodeLIKE(TBL, vCriteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                QueryResult = DTA.FillBy_NoGRNLIKE(TBL, vCriteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                QueryResult = DTA.FillBy_NoStockInPurchaseLIKE(TBL, vCriteria + "%", IsQuick)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, IsQuick, vCriteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, IsQuick, vCriteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns("Date").DefaultCellStyle = ij
            End If

            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > 10000000 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Ticks / TimeSpan.TicksPerSecond) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + " miliseconds"
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

    Public Function GetTanggalTransaksi(ByVal vPrimaryPR As Double) As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            If DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggalTransaksi(vPrimaryPR).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = True, _
    Optional ByVal vStrictEdit As Boolean = True, Optional ByVal vQuickPR As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
        Dim NoStockOut As String = ""
        Dim ObjAOP As New ClsTransaksiAOP
        Dim NoAOP As String = ""
        Dim TanggalPR As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosing
        Dim Jurnal As String = ""
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturn
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Try
            'yang boleh menghapus adalah yang menciptakan
            '0. strict deleting
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be edited by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            'purchase return tidak boleh diedit bila:
            '0. sudah closing
            TanggalPR = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalPR Then
                MsgBox("The period " + Format(TanggalPR, "MMM yyyy") + " has been closed, you cannot edit this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. Jika sudah ada AOP yang merujuk ke purchase return ini
            'untuk AOP, berikan pesan no AOP dan no stock out purchase return
            NoAOP = ObjAOP.GetNoAOPBy_PrimaryPurchaseReturn(Me.mPrimary)
            If NoAOP <> "" Then
                If TampilkanPesan = True Then
                    NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                    MsgBox("This purchase return is refer by AOP #" + NoAOP + _
                    " and stock out #" + NoStockOut + ". Please delete the transaction to edit this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. jika sudah ada purchase return receivement journal
            Jurnal = ObjJurnal.Get_PurchaseReturn_Jurnal(Me.mPrimary)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This data is refered by purchase return receivement jounal: " + Jurnal + _
                    ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '4. sudah ada stock out purchase return yang merujuk ke purchase return ini
            If vQuickPR = False Then
                If DTA.ScalarQuery_GetJmlStockOut_PR(TipeTrans, Me.mPrimary).Value > 0 Then
                    If TampilkanPesan = True Then
                        NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                        If MsgBox("This purchase return is refer by stock out #" + NoStockOut + _
                        ". To continue editting this data, you have to delete stock out transaction, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
                            Me.mPrimaryStockOutPurchaseReturn = DTA.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                            ObjStockInOut.CreatedBy = ObjStockInOut.GetCreatedBy(Me.mPrimaryStockOutPurchaseReturn)
                            ObjStockInOut.Primary = Me.mPrimaryStockOutPurchaseReturn
                            ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                            If ObjStockInOut.Hapus(Me.mPrimaryStockOutPurchaseReturn, Me.mCreatedBy, TblItemCustom) = False Then
                                ObjStockInOut = Nothing
                                Return False
                            End If
                            Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                            ObjStockInOut = Nothing
                        Else
                            Return False
                        End If
                    End If
                End If
            End If

            If ObjCustom.ApakahBolehEdit_Custom(Me.mPrimary, TampilkanPesan) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjAOP = Nothing
            ObjJurnal = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vUser As String, Optional ByVal TampilkanPesan As Boolean = False, _
    Optional ByVal vStrictEdit As Boolean = True, Optional ByVal vQuickPR As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
        Dim NoStockOut As String = ""
        Dim ObjAOP As New ClsTransaksiAOP
        Dim NoAOP As String = ""
        Dim TanggalPR As Date = #1/1/1950#
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCL As New ClsTransaksiMonthlyClosing
        Dim ObjJurnal As New AccModule.ClsTransaksiJournal
        Dim Jurnal As String = ""
        Dim ObjCustom As New CustomSMESS.ClsTransaksiPurchaseReturn
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Try
            'yang boleh menghapus adalah yang menciptakan
            '0. strict deleting
            If vStrictEdit = True Then
                If vUser <> Me.mCreatedBy Then
                    If TampilkanPesan = True Then
                        MsgBox("This transaction can only be deleted by: " + Me.mCreatedBy, MsgBoxStyle.Information, _
                        "Information")
                    End If

                    Return False
                End If
            End If

            'purchase return tidak boleh diedit bila:
            'sudah clsoing
            TanggalPR = Me.GetTanggalTransaksi(Me.mPrimary)
            If ObjCD.GetLasDateOfTheMonth(ObjCL.GetLastClosingPeriod()) >= TanggalPR Then
                MsgBox("The period " + Format(TanggalPR, "MMM yyyy") + " has been closed, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            '1. Jika sudah ada AOP yang merujuk ke purchase return ini
            'untuk AOP, berikan pesan no AOP dan no stock out purchase return
            NoAOP = ObjAOP.GetNoAOPBy_PrimaryPurchaseReturn(Me.mPrimary)
            If NoAOP <> "" Then
                If TampilkanPesan = True Then
                    NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                    MsgBox("This purchase return is refer by AOP #" + NoAOP + _
                    ". Please delete the transaction to delete this data.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. jika sudah ada purchase retun receivement jurnal
            Jurnal = ObjJurnal.Get_PurchaseReturn_Jurnal(Me.mPrimary)
            If Jurnal <> "" Then
                If TampilkanPesan = True Then
                    MsgBox("This data is refered by purchase return receivement jounal: " + Jurnal + _
                    ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            If vQuickPR = False Then
                '4. sudah ada stock out purchase return yang merujuk ke purchase return ini
                'If DTA.ScalarQuery_GetJmlStockOut_PR(TipeTrans, Me.mPrimary).Value > 0 Then
                '    'If TampilkanPesan = True Then
                '    'NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                '    'MsgBox("This purchase return is refer by stock out #" + NoStockOut + _
                '    '". Please delete the transaction to delete this data.", MsgBoxStyle.Information, "Information")
                '    'End If

                '    Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
                '    ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                '    ObjStockInOut.CreatedBy = ObjStockInOut.GetCreatedBy(Me.mPrimaryStockOutPurchaseReturn)
                '    ObjStockInOut.Primary = Me.mPrimaryStockOutPurchaseReturn
                '    ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                '    If ObjStockInOut.ApakahBolehHapus(Me.mCreatedBy) = False Then
                '        ObjStockInOut = Nothing
                '        Return False
                '    End If
                '    ObjStockInOut = Nothing
                'End If

                '5. sudah ada stock out purchase return yang merujuk ke purchase return ini
                If DTA.ScalarQuery_GetJmlStockOut_PR(TipeTrans, Me.mPrimary).Value > 0 Then
                    If TampilkanPesan = True Then
                        NoStockOut = DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                        If MsgBox("This purchase return is refer by stock out #" + NoStockOut + _
                        ". To continue deleting this data, you have to delete stock out transaction, continue?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                            Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
                            ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                            ObjStockInOut.CreatedBy = ObjStockInOut.GetCreatedBy(Me.mPrimaryStockOutPurchaseReturn)
                            ObjStockInOut.Primary = Me.mPrimaryStockOutPurchaseReturn
                            If ObjStockInOut.ApakahBolehHapus(Me.mCreatedBy) = True Then
                                If ObjStockInOut.Hapus(Me.mPrimaryStockOutPurchaseReturn, Me.mCreatedBy, TblItemCustom) = False Then
                                    ObjStockInOut = Nothing
                                    Return False
                                Else
                                    Return True
                                End If
                            Else
                                Return False
                            End If
                            Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                            ObjStockInOut = Nothing
                        Else
                            Return False
                        End If
                    End If
                End If
            End If

            '6. cek berkaitan dengan minus pada purchase order balance 
            Return Me.CekPOB(Me.mPrimary)

            If ObjCustom.ApakahBolehHapus_Custom(Me.mPrimary, TampilkanPesan) = False Then
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjAOP = Nothing
            ObjJurnal = Nothing
            ObjCustom = Nothing
        End Try
    End Function

    Public Function GetNoTransaksiStockOut(ByVal vPrimaryPurchaseReturn As Double) As String
        Dim DTA As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
        Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = _
        Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
        Dim NoStockOut As String = ""

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, _
            vPrimaryPurchaseReturn)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoStockOutPurchaseReturn(TipeTrans, _
                vPrimaryPurchaseReturn)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDgvItem As DataGridView, _
    ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, ByVal vPrimaryStockInPurchase As Double, _
    ByVal vPrimaryGRN As Double, ByVal vPrimaryPR As Double, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item3TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item3DataTable
        Dim rowInOutLama As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item3Row

        Dim DTAMasterProductQty As New Core_Module.dseMasterProductTableAdapters.tblmaster_productqtyTableAdapter
        DTAMasterProductQty.ClearBeforeFill = False
        Dim TBLMasterProductQty As New Core_Module.dseMasterProduct.tblmaster_productqtyDataTable

        Dim rowDGV As DataGridViewRow
        Dim BatchNo As String
        Dim PrimaryProduct, PrimaryLocation, Sec_Qty, Qty As Decimal
        Dim ExpiryDate As Date

        Dim ListPrimaryProduct As New Queue(Of Double)
        Dim rowProduct As Core_Module.dseMasterProduct.tblmaster_productqtyRow
        Dim Dummy As String = ""

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'simulasikan stock minus untuk setiap item transaksi
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.Fill(TBL, Me.mPrimary)
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                        -1, .Primary_Product, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'load data item baru
            For Each rowDGV In vDgvItem.Rows
                BatchNo = rowDGV.Cells(vFrm.ColBatchNo.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiryDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                DTAMasterProductQty.FillBy_Parameter(TBLMasterProductQty, _
                Sec_Qty, PrimaryProduct, BatchNo, ExpiryDate)
            Next

            'simulasikan untuk stock out
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowInOutLama In TBL.Rows
                    With rowInOutLama
                        Me.Koreksi4HapusOut(TBLMasterProductQty, _
                        .Primary_Product, .Qty, -1, _
                        .primary_location, .Batch_No, .Expiry_Date)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDgvItem.Rows
                BatchNo = rowDGV.Cells(vFrm.ColBatchNo.Index).Value
                PrimaryLocation = rowDGV.Cells(vFrm.ColPrimaryLocation.Index).Value
                Sec_Qty = "-1"
                If IsDate(rowDGV.Cells(vFrm.ColExpiryDate.Index).Value) = False Then
                    ExpiryDate = #1/1/1950#
                Else
                    ExpiryDate = CDate(rowDGV.Cells(vFrm.ColExpiryDate.Index).Value)
                End If
                PrimaryProduct = rowDGV.Cells(vFrm.ColPrimaryCodeNo.Index).Value
                Qty = rowDGV.Cells(vFrm.ColTotalQty.Index).Value

                Me.Koreksi4SimpanOut(TBLMasterProductQty, _
                PrimaryProduct, Qty, Sec_Qty, PrimaryLocation, _
                BatchNo, ExpiryDate)
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowProduct In TBLMasterProductQty.Rows
                If rowProduct.Qty < 0 Then
                    If Me.ObjProduct.ApakahGoods(rowProduct.Primary_Product) = True Then
                        ListPrimaryProduct.Enqueue(rowProduct.Primary_Product)
                    End If
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    Else
                        Dummy = Dummy + ", " + ObjProduct.GetProductCode(ListPrimaryProduct.Dequeue)
                    End If
                End While
                ObjProduct = Nothing

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'simlasikan GRN minus untuk setiap item transaksi
            Return Me.ApakahDataBenar(vBaru, vDgvItem, vFrm.ColPrimaryCodeNo.Index, vFrm.ColPrimaryPO.Index, vFrm.ColTotalQty.Index, _
            vPrimaryStockInPurchase, vPrimaryGRN, vPrimaryPR)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAMasterProductQty.Dispose()
            TBLMasterProductQty.Dispose()
            ListPrimaryProduct = Nothing
        End Try
    End Function

    Private Sub GetListReturnBalance_Default(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, ByVal vPrimaryStockIn As Double, _
    ByVal vPrimaryGRN As Double)
        Dim DTA_Dummy As New DseTransaksiPurchaseReturnTableAdapters.PR_OBTableAdapter
        Dim TBL_Dummy As New DseTransaksiPurchaseReturn.PR_OBDataTable
        Dim row_Dummy, row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2TableAdapter
        Dim TBL1 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2DataTable
        Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2Row
        Dim Filter As String = ""

        Try
            'isikan data return balance pada dummy: total stock in - total yang sudah direturn
            'total stock in
            DTA_Dummy.Fill(TBL_Dummy, vPrimaryStockIn)
            'total return
            DTA1.Fill(TBL1, vPrimaryGRN)
            'kurangkan
            For Each row1 In TBL1.Rows
                If row1.Primary_PO <> -1 Then
                    Filter = "[Primary Product] = " + row1.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row1.Primary_PO.ToString
                    TBL_Dummy.DefaultView.RowFilter = Filter

                    While TBL_Dummy.DefaultView.Count > 0 And row1.Total_Qty > 0
                        row_Dummy = TBL_Dummy.DefaultView(0).Row
                        If row_Dummy.Qty > row1.Total_Qty Then
                            row_Dummy.Qty -= row1.Total_Qty
                            row1.Total_Qty = 0
                        Else
                            row1.Total_Qty -= row_Dummy.Qty
                            TBL_Dummy.Rows.Remove(row_Dummy)
                        End If
                    End While
                End If
            Next

            'pindahkan ke tbl
            For Each row_Dummy In TBL_Dummy.Rows
                Filter = "[Primary Product] = " + row_Dummy.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row_Dummy.Primary_PO_Purchase.ToString
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    row.Qty += row_Dummy.Qty
                Else
                    row = vTBL.NewPR_OBRow

                    With row_Dummy
                        row.PO_Number = .PO_Number
                        row.Primary_PO_Purchase = .Primary_PO_Purchase
                        row.Primary_Product = .Primary_Product
                        row.Product_Code = .Product_Code
                        row.Qty = .Qty
                    End With

                    vTBL.Rows.Add(row)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Dummy.Dispose()
            TBL_Dummy.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub GetListReturnBalance_Multi(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, ByVal vPrimaryStockIn As Double, _
    ByVal vPrimaryGRN As Double)
        Dim DTA_Dummy As New DseTransaksiPurchaseReturnTableAdapters.PR_OB_MultiTableAdapter
        Dim TBL_Dummy As New DseTransaksiPurchaseReturn.PR_OB_MultiDataTable
        Dim row_Dummy, row As DseTransaksiPurchaseReturn.PR_OB_MultiRow
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2_MultiTableAdapter
        Dim TBL1 As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2_MultiDataTable
        Dim row1 As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2_MultiRow
        Dim Filter As String = ""

        Try
            'isikan data return balance pada dummy: total stock in - total yang sudah direturn
            'total stock in
            DTA_Dummy.Fill(TBL_Dummy, vPrimaryStockIn)
            'total return
            DTA1.Fill(TBL1, vPrimaryGRN)
            'kurangkan
            For Each row1 In TBL1.Rows
                If row1.Primary_PO <> -1 Then
                    Filter = "[Primary Product] = " + row1.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row1.Primary_PO.ToString
                    TBL_Dummy.DefaultView.RowFilter = Filter

                    While TBL_Dummy.DefaultView.Count > 0 And row1.Total_Qty_Multi > 0
                        row_Dummy = TBL_Dummy.DefaultView(0).Row
                        If row_Dummy.Qty_Multi > row1.Total_Qty_Multi Then
                            row_Dummy.Qty_Multi -= row1.Total_Qty_Multi
                            row1.Total_Qty_Multi = 0
                        Else
                            row1.Total_Qty_Multi -= row_Dummy.Qty_Multi
                            TBL_Dummy.Rows.Remove(row_Dummy)
                        End If
                    End While
                End If
            Next

            'pindahkan ke tbl
            For Each row_Dummy In TBL_Dummy.Rows
                Filter = "[Primary Product] = " + row_Dummy.Primary_Product.ToString + " AND [Primary PO Purchase] = " + row_Dummy.Primary_PO_Purchase.ToString
                vTBL.DefaultView.RowFilter = Filter

                If vTBL.DefaultView.Count > 0 Then
                    row = vTBL.DefaultView(0).Row
                    row.Qty_Multi += row_Dummy.Qty_Multi
                Else
                    row = vTBL.NewPR_OB_MultiRow

                    With row_Dummy
                        row.PO_Number = .PO_Number
                        row.Primary_PO_Purchase = .Primary_PO_Purchase
                        row.Primary_Product = .Primary_Product
                        row.Product_Code = .Product_Code
                        row.Qty_Multi = .Qty_Multi
                    End With

                    vTBL.Rows.Add(row)
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Dummy.Dispose()
            TBL_Dummy.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
        End Try
    End Sub

    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, ByVal vIndexPrimaryCodeNo As Integer, _
    ByVal vIndexPrimaryPO As Integer, ByVal vIndexQtyMulti As Integer, ByVal vPrimaryStockInPurchase As Double, ByVal vPrimaryGRN As Double, _
    ByVal vPrimaryPR As Double, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Dim DTA_Lama As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item2_MultiTableAdapter
        Dim TBL_Lama As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2_MultiDataTable
        Dim row_Lama As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item2_MultiRow

        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.PR_OB_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.PR_OB_MultiDataTable
        Dim row1 As DseTransaksiPurchaseReturn.PR_OB_MultiRow

        Dim rowDGV As DataGridViewRow
        Dim ListPrimaryProduct As New Queue(Of String)

        Dim Dummy As String = ""
        Dim PrimaryGRNLama, PrimaryStockInPurchaseLama, PrimaryProduct, PrimaryPO As Double
        Dim ObjGRN As New ClsTransaksiGRN
        Dim ObjPO As New ClsTransaksiPO
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim Qty As Decimal = 0

        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'simulasikan return balance minus
            PrimaryGRNLama = Me.GetPrimaryGRN(vPrimaryPR)
            PrimaryStockInPurchaseLama = ObjGRN.GetPrimaryStockIn(PrimaryGRNLama)

            If vBaru = False Then
                'load data return dari pr lama
                If PrimaryGRNLama <> -1 Then
                    Me.GetListReturnBalance_Multi(TBL, PrimaryStockInPurchaseLama, PrimaryGRNLama)
                End If
                'isikan data lama, untuk ditambahkan ke dalam return balance
                DTA_Lama.FillBy_PrimaryPR(TBL_Lama, vPrimaryGRN, Me.mPrimary)
            End If

            'load data return balance untuk pr baru, jika ternyata stock in purchasenya tidak sama antara yang lama dan yang baru
            If vPrimaryGRN <> -1 Then
                If PrimaryGRNLama <> vPrimaryGRN Then
                    Me.GetListReturnBalance_Multi(TBL, vPrimaryStockInPurchase, vPrimaryGRN)
                End If
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                If vIndexPrimaryPO <> -1 Then
                    PrimaryProduct = rowDGV.Cells(vIndexPrimaryCodeNo).Value
                    PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value
                    Qty = rowDGV.Cells(vIndexQtyMulti).Value

                    If PrimaryPO <> -1 Then
                        Me.Koreksi4SimpanOut_ReturnBalance_Multi(TBL, _
                        PrimaryProduct, Qty, PrimaryPO, ObjPO, ObjProduct)
                    End If
                End If
            Next

            'simulasi hapus data item lama
            If TBL_Lama.Rows.Count > 0 Then
                For Each row_Lama In TBL_Lama.Rows
                    With row_Lama
                        Me.Koreksi4HapusOut_ReturnBalance_Multi(TBL, row_Lama.Primary_Product, Qty, row_Lama.Primary_PO, ObjProduct, ObjPO)
                    End With
                Next
            End If

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each row1 In TBL.Rows
                If row1.Qty_Multi < 0 Then
                    ListPrimaryProduct.Enqueue(row1.Product_Code + "(PO : " + row1.PO_Number + ")")
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = ListPrimaryProduct.Dequeue
                    Else
                        Dummy = Dummy + ", " + ListPrimaryProduct.Dequeue
                    End If
                End While

                MsgBox("The following product(s) return balance will be minus: " + Dummy + ", process cancelled.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'cek pob
            Return CekPOB_Multi(vBaru, vDGVItem, vIndexPrimaryCodeNo, vIndexQtyMulti, vIndexPrimaryPO)

            'kalo sampa sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTA_Lama.Dispose()
            TBL_Lama.Dispose()
            ListPrimaryProduct = Nothing
            ObjGRN = Nothing
            ObjPO = Nothing
            ObjProduct = Nothing
        End Try
    End Function

    Private Function CekPOB(ByVal vPrimaryPR As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item4_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4_MultiDataTable
        Dim rowPR_Lama As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4_MultiRow

        Dim DTAPOB As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_po_item_MultiTableAdapter
        DTAPOB.ClearBeforeFill = False
        Dim TBLPOB As New DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable
        Dim rowPOB As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow

        Dim ListPO As New List(Of Double)
        Dim ListPrimaryProduct As New Queue(Of String)
        Dim Dummy As String = ""

        Dim DtaSip As New DseTransaksiPurchaseReturnTableAdapters.DataTable7TableAdapter
        Dim TblSip As New DseTransaksiPurchaseReturn.DataTable7DataTable
        Dim rowSip As DseTransaksiPurchaseReturn.DataTable7Row

        Try
            DTA.Fill(TBL, vPrimaryPR)
            'load data item lama
            DTA.ClearBeforeFill = False
            For Each rowPR_Lama In TBL.Rows
                If rowPR_Lama.Primary_PO <> -1 Then
                    DTAPOB.Fill(TBLPOB, rowPR_Lama.Primary_PO, rowPR_Lama.Primary_Product)
                    If ListPO.Contains(rowPR_Lama.Primary_PO) = False Then
                        ListPO.Add(rowPR_Lama.Primary_PO)
                    End If
                End If
            Next

            'For Each rowPR_Lama In TBL.Rows
            '    If ListPO.Contains(rowPR_Lama.Primary_PO) = False And rowPR_Lama.Primary_PO <> -1 Then
            '        DTAPOB.Fill(TBLPOB, rowPR_Lama.Primary_PO, rowPR_Lama.Primary_Product)
            '        ListPO.Add(rowPR_Lama.Primary_PO)
            '    End If
            'Next

            If ListPO.Count > 0 Then
                ListPO.Sort()
                DtaSip.FillBy_PrimaryPOMinMax(TblSip, ListPO.Item(0), ListPO.Item(ListPO.Count - 1))
            End If

            'simulasikan untuk stock in hapus
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowPR_Lama In TBL.Rows
                    With rowPR_Lama
                        Me.Koreksi4HapusIn_POB_Multi(TBLPOB, _
                        .Primary_Product, .Qty_Multi, .Primary_PO)
                    End With
                Next
            End If

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowPOB In TBLPOB.Rows
                'jika pob minus, cek apakah minus karena ada tolerance
                TblSip.DefaultView.RowFilter = "[Primary PO] = " & rowPOB.Primary_PO.ToString & " AND [Primary Product] = " & rowPOB.Primary_Product
                If TblSip.DefaultView.Count > 0 Then
                    rowSip = TblSip.DefaultView(0).Row
                    'jika qty stock in purchase lebih besar dari qty po maka ada tolerance
                    If rowSip.Qty_Multi < rowSip.Total_Qty_Multi Then
                        rowPOB.POB_Multi = 0
                    End If
                End If
                If rowPOB.POB_Multi < 0 Then
                    ListPrimaryProduct.Enqueue(rowPOB.Primary_Product & "-" & rowPOB.Primary_PO)
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                Dim ObjPO As New ClsTransaksiPO

                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                    Else
                        Dummy = Dummy + ", " + GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                    End If
                End While
                ObjProduct = Nothing
                ObjPO = Nothing

                MsgBox("The purchase order balance for the following product(s) will be minus: " + Dummy + ", process cancelled.", _
               MsgBoxStyle.Information, "Information")

                Return False
            End If

            'kalo mpe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAPOB.Dispose()
            TBLPOB.Dispose()
            ListPrimaryProduct = Nothing
            ListPO = Nothing
            DtaSip = Nothing
            TblSip = Nothing
        End Try
    End Function

    Private Function CekPOB_Default(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, _
    ByVal vIndexPrimaryProduct As Integer, ByVal vIndexQty As Integer, ByVal vIndexPrimaryPO As Integer) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item4TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4DataTable
        Dim rowPR_Lama As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4Row

        Dim DTAPOB As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_po_itemTableAdapter
        DTAPOB.ClearBeforeFill = False
        Dim TBLPOB As New DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable
        Dim rowPOB As DseTransaksiPurchaseReturn.tbltrans_po_itemRow

        Dim rowDGV As DataGridViewRow
        Dim PrimaryProduct, PrimaryPO As Double
        Dim Qty As Decimal
        Dim ListPO As New List(Of Double)
        Dim ListPrimaryProduct As New Queue(Of String)
        Dim Dummy As String = ""

        Try
            DTA.ClearBeforeFill = False
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.Fill(TBL, Me.mPrimary)
                For Each rowPR_Lama In TBL.Rows
                    With rowPR_Lama
                        If .Primary_PO <> -1 Then
                            DTAPOB.Fill(TBLPOB, .Primary_PO)
                            If ListPO.Contains(.Primary_PO) = False Then
                                ListPO.Add(.Primary_PO)
                            End If
                        End If
                    End With
                Next
            End If

            'load data item baru
            For Each rowDGV In vDGVItem.Rows
                PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value
                If PrimaryPO <> -1 Then
                    DTAPOB.Fill(TBLPOB, PrimaryPO)
                    If ListPO.Contains(PrimaryPO) = False Then
                        ListPO.Add(PrimaryPO)
                    End If
                End If
            Next

            'If vBaru = False Then
            '    'load database master product qty berdasarkan data product lama dan data product baru
            '    TBL.Rows.Clear()
            '    DTA.Fill(TBL, Me.mPrimary)
            '    For Each rowPR_Lama In TBL.Rows
            '        With rowPR_Lama
            '            If ListPO.Contains(.Primary_PO) = False Then
            '                DTAPOB.Fill(TBLPOB, .Primary_PO)
            '                ListPO.Add(.Primary_PO)
            '            End If
            '        End With
            '    Next
            'End If

            ''load data item baru
            'For Each rowDGV In vDGVItem.Rows
            '    PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value
            '    If ListPO.Contains(PrimaryPO) = False And PrimaryPO <> -1 Then
            '        DTAPOB.Fill(TBLPOB, PrimaryPO)
            '        ListPO.Add(PrimaryPO)
            '    End If
            'Next

            'simulasikan untuk stock in 
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowPR_Lama In TBL.Rows
                    With rowPR_Lama
                        Me.Koreksi4HapusIn_POB_Default(TBLPOB, _
                        .Primary_Product, .Qty, .Primary_PO)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(vIndexPrimaryProduct).Value
                Qty = rowDGV.Cells(vIndexQty).Value '
                PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value

                If PrimaryPO <> -1 Then
                    Me.Koreksi4SimpanIn_POB_Default(TBLPOB, _
                    PrimaryProduct, Qty, PrimaryPO)
                End If
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            For Each rowPOB In TBLPOB.Rows
                If rowPOB.POB < 0 Then
                    ListPrimaryProduct.Enqueue(rowPOB.Primary_Product & "-" & rowPOB.Primary_PO)
                End If
            Next
            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                Dim ObjPO As New ClsTransaksiPO

                While ListPrimaryProduct.Count > 0
                    If Dummy = "" Then
                        Dummy = GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                    Else
                        Dummy = Dummy + ", " + GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                    End If
                End While
                ObjProduct = Nothing
                ObjPO = Nothing

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
               MsgBoxStyle.Information, "Information")

                Return False
            End If

            'kalo mpe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAPOB.Dispose()
            TBLPOB.Dispose()
            ListPrimaryProduct = Nothing
            ListPO = Nothing
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function CekPOB_Multi(ByVal vBaru As Boolean, ByVal vDGVItem As DataGridView, _
    ByVal vIndexPrimaryProduct As Integer, ByVal vIndexQty As Integer, ByVal vIndexPrimaryPO As Integer) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item4_MultiTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4_MultiDataTable
        Dim rowPR_Lama As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item4_MultiRow

        Dim DTAPOB As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_po_item_MultiTableAdapter
        DTAPOB.ClearBeforeFill = False
        Dim TBLPOB As New DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable
        Dim rowPOB As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow

        Dim rowDGV As DataGridViewRow
        Dim PrimaryProduct, PrimaryPO As Double
        Dim Qty As Decimal
        Dim ListPO As New List(Of Double)
        Dim ListPrimaryProduct As New Queue(Of String)
        Dim Dummy As String = ""

        Dim DtaSip As New DseTransaksiPurchaseReturnTableAdapters.DataTable7TableAdapter
        Dim TblSip As New DseTransaksiPurchaseReturn.DataTable7DataTable
        Dim rowSip As DseTransaksiPurchaseReturn.DataTable7Row

        Try
            DTA.ClearBeforeFill = False
            If vBaru = False Then
                'load database master product qty berdasarkan data product lama dan data product baru
                TBL.Rows.Clear()
                DTA.Fill(TBL, Me.mPrimary)
                For Each rowPR_Lama In TBL.Rows
                    With rowPR_Lama
                        If .Primary_PO <> -1 Then
                            DTAPOB.Fill(TBLPOB, .Primary_PO, .Primary_Product)
                            If ListPO.Contains(.Primary_PO) = False Then
                                ListPO.Add(.Primary_PO)
                            End If
                        End If
                    End With
                Next
            End If

            'load data item baru
            For Each rowDGV In vDGVItem.Rows
                PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value
                PrimaryProduct = rowDGV.Cells(vIndexPrimaryProduct).Value
                If PrimaryPO <> -1 Then
                    DTAPOB.Fill(TBLPOB, PrimaryPO, PrimaryProduct)
                    If ListPO.Contains(PrimaryPO) = False Then
                        ListPO.Add(PrimaryPO)
                    End If
                End If
            Next

            'If vBaru = False Then
            '    'load database master product qty berdasarkan data product lama dan data product baru
            '    TBL.Rows.Clear()
            '    DTA.Fill(TBL, Me.mPrimary)
            '    For Each rowPR_Lama In TBL.Rows
            '        With rowPR_Lama
            '            If ListPO.Contains(.Primary_PO) = False Then
            '                DTAPOB.Fill(TBLPOB, .Primary_PO, .Primary_Product)
            '                ListPO.Add(.Primary_PO)
            '            End If
            '        End With
            '    Next
            'End If

            ''load data item baru
            'For Each rowDGV In vDGVItem.Rows
            '    PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value
            '    PrimaryProduct = rowDGV.Cells(vIndexPrimaryProduct).Value
            '    If ListPO.Contains(PrimaryPO) = False And PrimaryPO <> -1 Then
            '        DTAPOB.Fill(TBLPOB, PrimaryPO, PrimaryProduct )
            '        ListPO.Add(PrimaryPO)
            '    End If
            'Next

            If ListPO.Count > 0 Then
                ListPO.Sort()
                DtaSip.FillBy_PrimaryPOMinMax(TblSip, ListPO.Item(0), ListPO.Item(ListPO.Count - 1))
            End If

            'simulasikan untuk stock in 
            'simulasi hapus data item lama
            If TBL.Rows.Count > 0 Then
                For Each rowPR_Lama In TBL.Rows
                    With rowPR_Lama
                        Me.Koreksi4HapusIn_POB_Multi(TBLPOB, _
                        .Primary_Product, .Qty_Multi, .Primary_PO)
                    End With
                Next
            End If

            'simulasi simpan item baru
            For Each rowDGV In vDGVItem.Rows
                PrimaryProduct = rowDGV.Cells(vIndexPrimaryProduct).Value
                Qty = rowDGV.Cells(vIndexQty).Value '
                PrimaryPO = rowDGV.Cells(vIndexPrimaryPO).Value

                If PrimaryPO <> -1 Then
                    Me.Koreksi4SimpanIn_POB_Multi(TBLPOB, _
                    PrimaryProduct, Qty, PrimaryPO, TblSip)
                End If
            Next

            'cek apakah ada qty yang minus setelah dilakukan simulasi, jika ada
            'maka return false
            ListPrimaryProduct.Clear()
            If TBLPOB.Rows.Count > 0 Then
                For Each rowPOB In TBLPOB.Rows
                    'jika pob minus, cek apakah minus karena ada tolerance
                    TblSip.DefaultView.RowFilter = "[Primary PO] = " & rowPOB.Primary_PO.ToString & " AND [Primary Product] = " & rowPOB.Primary_Product
                    If TblSip.DefaultView.Count > 0 Then
                        rowSip = TblSip.DefaultView(0).Row
                        'jika qty stock in purchase lebih besar dari qty po maka ada tolerance
                        If rowSip.Qty_Multi < rowSip.Total_Qty_Multi Then
                            rowPOB.POB_Multi = 0
                        End If
                    End If
                    If rowPOB.POB_Multi < 0 Then
                        ListPrimaryProduct.Enqueue(rowPOB.Primary_Product & "-" & rowPOB.Primary_PO)
                    End If
                Next
            End If

            If ListPrimaryProduct.Count > 0 Then
                Dim ObjProduct As New Core_Module.clsMasterProduct
                Dim ObjPO As New ClsTransaksiPO

                While ListPrimaryProduct.Count > 0
                    If IsNumeric(ListPrimaryProduct.Dequeue) Then
                        If Dummy = "" Then
                            Dummy = GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                        Else
                            Dummy = Dummy + ", " + GetProductCode_PONumber(ListPrimaryProduct.Dequeue, ObjProduct, ObjPO)
                        End If
                    End If
                End While
                ObjProduct = Nothing
                ObjPO = Nothing

                MsgBox("The following product(s) will be minus: " + Dummy + ", process cancelled.", _
               MsgBoxStyle.Information, "Information")

                Return False
            End If

            'kalo mpe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            DTAPOB.Dispose()
            TBLPOB.Dispose()
            ListPrimaryProduct = Nothing
            ListPO = Nothing
            DtaSip = Nothing
            TblSip = Nothing
        End Try
    End Function

    Private Function GetProductCode_PONumber(ByVal vStringToParse As String, _
    ByVal vObjProduct As Core_Module.clsMasterProduct, ByVal vObjPO As ClsTransaksiPO) As String
        Dim PrimaryProduct, PrimaryPO As Double
        Dim IndexAcuan As Integer

        Try
            IndexAcuan = vStringToParse.IndexOf("-")

            If IndexAcuan <> -1 Then
                PrimaryProduct = vStringToParse.Substring(0, IndexAcuan)
                PrimaryPO = vStringToParse.Substring(IndexAcuan + 1)
                Return vObjProduct.GetProductCode(PrimaryProduct) + "(" + vObjPO.GetNoTransaksiPO(PrimaryPO) + ")"
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function Koreksi4SimpanIn_POB_Default(ByVal vTBLPOB As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
   ByVal vPrimaryProduct As Double, _
   ByVal Qty As Decimal, ByVal vPrimaryPO As Double) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanIn_POB_Multi(ByVal vTBLPOB As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
   ByVal vPrimaryProduct As Double, _
   ByVal Qty As Decimal, ByVal vPrimaryPO As Double, ByVal TblSip As DseTransaksiPurchaseReturn.DataTable7DataTable) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0
        Dim rowSip As DseTransaksiPurchaseReturn.DataTable7Row
        Dim QtyBalance As Decimal = 0
        Try
            JmlLama = Me.GetJml_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO)
            'jika jmllama minus, cek apakah minus karena ada tolerance
            'If JmlLama < 0 Then
            '    TblSip.DefaultView.RowFilter = "[Primary PO] = " & vPrimaryPO.ToString & " AND [Primary Product] = " & vPrimaryProduct
            '    If TblSip.DefaultView.Count > 0 Then
            '        rowSip = TblSip.DefaultView(0).Row
            '        'jika qty sip lebih besar dari qty po maka ada tolerance
            '        If rowSip.Qty_Multi < rowSip.Total_Qty_Multi Then
            '            JmlLama = 0
            '        End If
            '    End If
            'End If
            
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO, Qty)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4HapusIn_POB_Default(ByVal vTBLPOB As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
   ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_POB_Default(vTBLPOB, vPrimaryProduct, vPrimaryPO, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Koreksi4HapusIn_POB_Multi(ByVal vTBLPOB As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
    ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_POB_Multi(vTBLPOB, vPrimaryProduct, vPrimaryPO, Qty * -1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Koreksi4SimpanOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, ByVal vPrimaryProduct As _
 Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty As Double = -1, _
 Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
 vBatchNo As String = "", Optional ByVal vExpiryDate As Date = #1/1/1950#) As Boolean
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

    Public Function Koreksi4SimpanOut_ReturnBalance_Default(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
    ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double, ByVal vObjPO As ClsTransaksiPO, _
    ByVal vObjProduct As Core_Module.clsMasterProduct) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO, Qty * -1, vObjPO, vObjProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function Koreksi4SimpanOut_ReturnBalance_Multi(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
    ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double, ByVal vObjPO As ClsTransaksiPO, _
    ByVal vObjProduct As Core_Module.clsMasterProduct) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO)

            If JmlLama <> 0 Then
                JmlBaru = JmlLama - Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO, Qty * -1, vObjPO, vObjProduct)
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

    Private Sub Hapus_Simulasi_POB_Default(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_itemRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Hapus_Simulasi_POB_Multi(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Hapus_Simulasi_ReturnBalance_Default(ByVal TBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
  ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double)
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString
            TBL.DefaultView.RowFilter = filter

            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Hapus_Simulasi_ReturnBalance_Multi(ByVal TBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double)
        Dim row As DseTransaksiPurchaseReturn.PR_OB_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString
            TBL.DefaultView.RowFilter = filter

            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                TBL.Rows.Remove(row)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

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

    Private Function GetJml_Simulasi_POB_Default(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_itemRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.POB
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJml_Simulasi_POB_Multi(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                Return row.POB_Multi
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Private Function GetJml_Simulasi_ReturnBalance_Default(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString
            vTBL.DefaultView.RowFilter = filter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.Qty
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Private Function GetJml_Simulasi_ReturnBalance_Multi(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double) As Decimal
        Dim row As DseTransaksiPurchaseReturn.PR_OB_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString
            vTBL.DefaultView.RowFilter = filter
            If vTBL.DefaultView.Count > 0 Then
                row = vTBL.DefaultView(0).Row
                Return row.Qty_Multi
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

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

    Private Sub SetJml_Simulasi_POB_Default(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
    ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
    ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_itemRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.POB = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub SetJml_Simulasi_POB_Multi(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
            ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO] = " + _
            vPrimaryPO.ToString
            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.POB_Multi = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub SetJml_Simulasi_ReturnBalance_Default(ByVal TBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
 ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.Qty = JmlBaru
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub SetJml_Simulasi_ReturnBalance_Multi(ByVal TBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.PR_OB_MultiRow
        Dim filter As String = ""

        Try
            filter = "[Primary Product] = " + vPrimaryProduct.ToString + " and [Primary PO Purchase] = " + _
            vPrimaryPO.ToString

            TBL.DefaultView.RowFilter = filter
            If TBL.DefaultView.Count > 0 Then
                row = TBL.DefaultView(0).Row
                row.Qty_Multi = JmlBaru
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

    Private Sub Insert_Simulasi_POB_Default(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_itemDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
   ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_itemRow

        Try
            row = TBL.Newtbltrans_po_itemRow

            row.Primary_Product = vPrimaryProduct
            row.Primary_PO = vPrimaryPO
            row.POB = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Insert_Simulasi_POB_Multi(ByVal TBL As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiDataTable, _
   ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
   ByVal JmlBaru As Decimal)
        Dim row As DseTransaksiPurchaseReturn.tbltrans_po_item_MultiRow

        Try
            row = TBL.Newtbltrans_po_item_MultiRow

            row.Primary_Product = vPrimaryProduct
            row.Primary_PO = vPrimaryPO
            row.POB_Multi = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub Insert_Simulasi_ReturnBalance_Default(ByVal TBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
  ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
  ByVal JmlBaru As Decimal, ByVal vObjPO As ClsTransaksiPO, ByVal vObjProduct As Core_Module.clsMasterProduct)
        Dim row As DseTransaksiPurchaseReturn.PR_OBRow

        Try
            row = TBL.NewPR_OBRow

            row.PO_Number = vObjPO.GetNoTransaksiPO(vPrimaryPO)
            row.Primary_PO_Purchase = vPrimaryPO
            row.Primary_Product = vPrimaryProduct
            row.Product_Code = vObjProduct.GetProductCode(vPrimaryProduct)
            row.Qty = JmlBaru

            TBL.Rows.Add(row)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''@rosadinnn
    ''@MUaT
    Private Sub Insert_Simulasi_ReturnBalance_Multi(ByVal TBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal vPrimaryPO As Double, _
            ByVal JmlBaru As Decimal, ByVal vObjPO As ClsTransaksiPO, ByVal vObjProduct As Core_Module.clsMasterProduct)
        Dim row As DseTransaksiPurchaseReturn.PR_OB_MultiRow

        Try
            row = TBL.NewPR_OB_MultiRow

            row.PO_Number = vObjPO.GetNoTransaksiPO(vPrimaryPO)
            row.Primary_PO_Purchase = vPrimaryPO
            row.Primary_Product = vPrimaryProduct
            row.Product_Code = vObjProduct.GetProductCode(vPrimaryProduct)
            row.Qty_Multi = JmlBaru

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

    Public Function Koreksi4HapusOut(ByVal TBLMasterProduct As Core_Module.dseMasterProduct.tblmaster_productqtyDataTable, _
  ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, Optional ByVal vSecondaryQty _
  As Double = -1, Optional ByVal vPrimaryLocation As Double = -1, Optional ByVal _
  vBatchNo As String = "", Optional ByVal vExpiryDate As Date = _
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
                    Me.SetJml_Simulasi(TBLMasterProduct, vPrimaryProduct, _
                    vSecondaryQty, vPrimaryLocation, vBatchNo, JmlBaru)
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

    Public Function Koreksi4HapusOut_ReturnBalance_Default(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OBDataTable, _
ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double, _
    ByVal vObjProduct As Core_Module.clsMasterProduct, ByVal vObjPO As ClsTransaksiPO) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_ReturnBalance_Default(vTBL, vPrimaryProduct, vPrimaryPO, Qty, vObjPO, vObjProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    ''@rosadinnn
    ''MUaT
    Public Function Koreksi4HapusOut_ReturnBalance_Multi(ByVal vTBL As DseTransaksiPurchaseReturn.PR_OB_MultiDataTable, _
            ByVal vPrimaryProduct As Double, ByVal Qty As Decimal, ByVal vPrimaryPO As Double, _
        ByVal vObjProduct As Core_Module.clsMasterProduct, ByVal vObjPO As ClsTransaksiPO) As Boolean
        Dim JmlLama As Decimal = 0
        Dim JmlBaru As Decimal = 0

        Try
            JmlLama = Me.GetJml_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO)
            If JmlLama <> 0 Then
                JmlBaru = JmlLama + Qty
                If JmlBaru = 0 Then
                    'hapus data lama
                    Me.Hapus_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO)
                Else
                    'update data jumlah
                    Me.SetJml_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, _
                    vPrimaryPO, JmlBaru)
                End If
            Else
                'masukkan data baru: tambah
                Me.Insert_Simulasi_ReturnBalance_Multi(vTBL, vPrimaryProduct, vPrimaryPO, Qty, vObjPO, vObjProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function


    Public Function ApakahBisaGenerateStockOutPurchaseReturn(ByVal vDGV As DataGridView, _
    ByVal i As FrmTransaksiPurchaseReturn) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tblmaster_productDataTable
        Dim rowDGV As DataGridViewRow
        Dim rowMaster As DseTransaksiPurchaseReturn.tblmaster_productRow

        Try
            DTA.ClearBeforeFill = False
            'isikan tabel dari master product
            For Each rowDGV In vDGV.Rows
                DTA.FillBy_PrimaryProduct(TBL, rowDGV.Cells(i.ColPrimaryCodeNo.Index).Value)
            Next
            'untuk setiap baris di dgv cek item-item yang dikontrol
            For Each rowMaster In TBL.Rows
                'cek apakah location control
                If rowMaster.Location_Controlled = 1 Then
                    Return False
                End If
                For Each rowDGV In vDGV.Rows
                    If rowDGV.Cells(i.ColCodeNo.Index).Value = rowMaster.Product_Code Then
                        'cek batch no controlled
                        If rowMaster.Batch_No_Controlled = 0 Then
                            If rowDGV.Cells(i.ColBatchNo.Index).Value <> "" Then
                                Return False
                            End If
                        Else
                            If rowDGV.Cells(i.ColBatchNo.Index).Value = "" Then
                                Return False
                            End If
                        End If
                        'expiry date controlled
                        If rowMaster.Expiry_Date_Controlled = True Then
                            If IsDate(rowDGV.Cells(i.ColExpiryDate.Index).Value) = False Then
                                Return False
                            End If
                        Else
                            If IsDate(rowDGV.Cells(i.ColExpiryDate.Index).Value) = True Then
                                Return False
                            End If
                        End If
                        'secondary qty
                        If rowMaster.Use_Secondary_Unit = 0 Then
                            If IsNumeric(rowDGV.Cells(i.ColSecQty.Index).Value) = True Then
                                Return False
                            End If
                        Else
                            If IsNumeric(rowDGV.Cells(i.ColSecQty.Index).Value) = False Then
                                Return False
                            End If
                        End If
                    End If
                Next
            Next

            'kalo sampe sini maka true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    'Public Function GenerateStockOutPurchaseReturn_Default(ByVal vDGVItem As DataGridView, _
    'ByVal i As FrmTransaksiPurchaseReturn) As Boolean
    '    Dim Obj As New Inventory_Module.clsTransaksiStockInOut
    '    Dim ObjGRN As New ClsTransaksiGRN
    '    Dim Frm As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
    '    Dim vDGVStockOut As DataGridView = Frm.dgvItem
    '    Dim rowSumber, rowTarget As DataGridViewRow
    '    Dim counter As Double = 0
    '    Dim SecQty As Double
    '    Dim ED As Date
    '    Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

    '    Try
    '        Obj.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
    '        Obj.TanggalTransaksi = Me.TanggalTransaksi
    '        Obj.Remarks = ""
    '        Obj.SupplierDoNo = Me.SupplierDONumber
    '        Obj.PrimarySupplier = ObjGRN.GetPrimarySupplier(Me.mPrimaryGRN)
    '        Obj.PrimaryPOPurchase = Me.mPrimary
    '        Obj.Primary = -1
    '        Obj.NoTransaksi = Obj.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn, _
    '        Me.TanggalTransaksi)

    '        'masukan data item ke dalam dgv
    '        vDGVStockOut.Rows.Clear()
    '        While counter < vDGVItem.Rows.Count
    '            rowSumber = vDGVItem.Rows(counter)
    '            vDGVStockOut.Rows.Add()
    '            rowTarget = vDGVStockOut.Rows(counter)
    '            With rowTarget
    '                If IsNumeric(rowSumber.Cells(i.ColSecQty.Index).Value) = True Then
    '                    SecQty = rowSumber.Cells(i.ColSecQty.Index).Value
    '                Else
    '                    SecQty = -1
    '                End If
    '                If IsDate(rowSumber.Cells(i.ColExpiryDate.Index).Value) = True Then
    '                    ED = rowSumber.Cells(i.ColExpiryDate.Index).Value
    '                Else
    '                    ED = #1/1/1950#
    '                End If

    '                .Cells(2).Value = Me.SupplierDONumber
    '                .Cells(5).Value = rowSumber.Cells(i.ColTotalQty.Index).Value
    '                .Cells(7).Value = rowSumber.Cells(i.ColSecQty.Index).Value
    '                .Cells(9).Value = rowSumber.Cells(i.ColBatchNo.Index).Value
    '                .Cells(10).Value = ED
    '                .Cells(13).Value = ""
    '                .Cells(14).Value = rowSumber.Cells(i.ColPrimaryPO.Index).Value
    '                .Cells(15).Value = rowSumber.Cells(i.ColPrimaryCodeNo.Index).Value
    '                .Cells(16).Value = -1
    '                .Cells(17).Value = rowSumber.Cells(i.ColPrimaryUnit.Index).Value
    '                .Cells(18).Value = rowSumber.Cells(i.ColPrimarySecUnit.Index).Value
    '                .Cells(19).Value = -1
    '            End With
    '            'next row
    '            counter += 1
    '        End While

    '        If Obj.SimpanBaru(vDGVStockOut, Me.CreatedBy) = True Then
    '            'Update primary stock out purchase return 
    '            Dim DTA1 As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
    '            Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn

    '            If DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary).HasValue = True Then
    '                Me.mPrimaryStockOutPurchaseReturn = DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
    '                DTA.UpdateQuery_PrimaryStockOutPurchaseReturn(Me.mPrimaryStockOutPurchaseReturn, Me.mPrimary)
    '                Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
    '            End If
    '            DTA1.Dispose()

    '            Obj.TransferJurnal(Obj.PrimarySupplier, Obj.Primary)

    '            Return True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Obj = Nothing
    '        Frm.Dispose()
    '        ObjGRN = Nothing
    '        DTA.Dispose()
    '    End Try
    'End Function

    '@rosadinnn
    '@MUaT
    Public Function GenerateStockOutPurchaseReturn_Multi(ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTransaksiPurchaseReturn) As Boolean
        Dim Obj As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjGRN As New ClsTransaksiGRN
        Dim Frm As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
        Dim vDGVStockOut As DataGridView = Frm.dgvItem
        Dim rowSumber, rowTarget As DataGridViewRow
        Dim counter As Double = 0
        Dim SecQty As Double
        Dim ED As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            Obj.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
            Obj.TanggalTransaksi = Me.TanggalTransaksi
            Obj.Remarks = ""
            Obj.SupplierDoNo = Me.SupplierDONumber
            Obj.PrimarySupplier = ObjGRN.GetPrimarySupplier(Me.mPrimaryGRN)
            Obj.PrimaryPOPurchase = Me.mPrimary
            Obj.Primary = -1
            Obj.NoTransaksi = Obj.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn, _
            Me.TanggalTransaksi)

            'masukan data item ke dalam dgv
            vDGVStockOut.Rows.Clear()
            While counter < vDGVItem.Rows.Count
                rowSumber = vDGVItem.Rows(counter)
                vDGVStockOut.Rows.Add()
                rowTarget = vDGVStockOut.Rows(counter)
                With rowTarget
                    If IsNumeric(rowSumber.Cells(i.ColSecQty.Index).Value) = True Then
                        SecQty = rowSumber.Cells(i.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If
                    If IsDate(rowSumber.Cells(i.ColExpiryDate.Index).Value) = True Then
                        ED = rowSumber.Cells(i.ColExpiryDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If

                    ''@rosadinnn
                    ''@MUaT
                    '.Cells(Frm.ColNo.Index).Value =
                    '.Cells(Frm.ColPONoDONo.Index).Value = 
                    .Cells("ColSuppDONo").Value = Me.SupplierDONumber
                    '.Cells(Frm.ColCodeNo.Index).Value = 
                    '.Cells(Frm.ColDescription.Index).Value = 
                    .Cells("ColQty").Value = rowSumber.Cells(i.ColTotalQty.Index).Value
                    '.Cells(Frm.ColUnit.Index).Value =
                    .Cells("ColSecondaryQty").Value = rowSumber.Cells(i.ColSecQty.Index).Value
                    '.Cells(Frm.ColSecondaryUnit.Index).Value = 
                    .Cells("ColBatch").Value = rowSumber.Cells(i.ColBatchNo.Index).Value
                    .Cells("ColExpiredDate").Value = ED
                    '.Cells(Frm.ColLocation.Index).Value = 
                    '.Cells(Frm.ColWareHouse.Index).Value =
                    .Cells("ColRemark").Value = ""
                    .Cells("ColPrimaryPOnoDoNo").Value = rowSumber.Cells(i.ColPrimaryPO.Index).Value
                    .Cells("ColPrimaryCodeNo").Value = rowSumber.Cells(i.ColPrimaryCodeNo.Index).Value
                    .Cells("ColPrimaryLocation").Value = -1
                    .Cells("ColPrimaryUnits").Value = rowSumber.Cells(i.ColPrimaryUnit.Index).Value
                    .Cells("ColPrimarySecondaryUnits").Value = rowSumber.Cells(i.ColPrimarySecUnit.Index).Value
                    .Cells("ColPrimaryPdO").Value = -1
                    .Cells("ColQtyMulti").Value = rowSumber.Cells(i.ColQtyMulti.Index).Value
                    '.Cells(Frm.ColUnitCodeMulti.Index).Value =
                    '.Cells(Frm.ColRatioDesc.Index).Value = 
                    .Cells("ColRatioReference").Value = rowSumber.Cells(i.ColRatioReference.Index).Value
                    .Cells("ColRatioDefault").Value = rowSumber.Cells(i.ColRatioDefault.Index).Value
                    .Cells("ColPrimaryUnitMulti").Value = rowSumber.Cells(i.ColPrimaryUnitMulti.Index).Value

                End With
                'next row
                counter += 1
            End While

            If Obj.SimpanBaru(vDGVStockOut, Me.CreatedBy, Frm.TblItemCustom) = True Then
                'Update primary stock out purchase return 
                Dim DTA1 As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
                Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn

                If DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary).HasValue = True Then
                    Me.mPrimaryStockOutPurchaseReturn = DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                    DTA.UpdateQuery_PrimaryStockOutPurchaseReturn(Me.mPrimaryStockOutPurchaseReturn, Me.mPrimary)
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
                End If
                DTA1.Dispose()

                'Obj.TransferJurnal(Obj.PrimarySupplier, Obj.Primary, -1)

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            Frm.Dispose()
            ObjGRN = Nothing
            DTA.Dispose()
        End Try
    End Function

    'Public Function GenerateStockOutPurchaseReturn_QuickPR_Default(ByVal vDGVItem As DataGridView, _
    '        ByVal i As FrmTransaksiPurchaseReturn_Quick, ByVal vNoTransaksiLama As String) As Boolean
    '    Dim Obj As New Inventory_Module.clsTransaksiStockInOut
    '    Dim ObjGRN As New ClsTransaksiGRN
    '    Dim Frm As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
    '    Dim vDGVStockOut As DataGridView = Frm.dgvItem
    '    Dim rowSumber, rowTarget As DataGridViewRow
    '    Dim counter As Double = 0
    '    Dim SecQty As Double
    '    Dim ED As Date
    '    Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

    '    Try
    '        Obj.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
    '        Obj.TanggalTransaksi = Me.TanggalTransaksi
    '        Obj.Remarks = ""
    '        Obj.SupplierDoNo = Me.SupplierDONumber
    '        Obj.PrimarySupplier = Me.mPrimarySupplier
    '        Obj.PrimaryPOPurchase = Me.mPrimary
    '        Obj.Primary = -1
    '        If vNoTransaksiLama <> "" Then
    '            Obj.NoTransaksi = vNoTransaksiLama
    '        Else
    '            Obj.NoTransaksi = Obj.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn, _
    '            Me.TanggalTransaksi)
    '        End If

    '        'masukan data item ke dalam dgv
    '        vDGVStockOut.Rows.Clear()
    '        While counter < vDGVItem.Rows.Count
    '            rowSumber = vDGVItem.Rows(counter)
    '            vDGVStockOut.Rows.Add()
    '            rowTarget = vDGVStockOut.Rows(counter)
    '            With rowTarget
    '                If IsNumeric(rowSumber.Cells(i.ColSecQty.Index).Value) = True Then
    '                    SecQty = rowSumber.Cells(i.ColSecQty.Index).Value
    '                Else
    '                    SecQty = -1
    '                End If
    '                If IsDate(rowSumber.Cells(i.ColExpiryDate.Index).Value) = True Then
    '                    ED = rowSumber.Cells(i.ColExpiryDate.Index).Value
    '                Else
    '                    ED = #1/1/1950#
    '                End If

    '                .Cells(2).Value = Me.SupplierDONumber
    '                .Cells(5).Value = rowSumber.Cells(i.ColTotalQty.Index).Value
    '                .Cells(7).Value = rowSumber.Cells(i.ColSecQty.Index).Value
    '                .Cells(9).Value = rowSumber.Cells(i.ColBatchNo.Index).Value
    '                .Cells(10).Value = ED
    '                .Cells(13).Value = ""
    '                .Cells(14).Value = rowSumber.Cells(i.ColPrimaryPO.Index).Value
    '                .Cells(15).Value = rowSumber.Cells(i.ColPrimaryCodeNo.Index).Value
    '                .Cells(16).Value = rowSumber.Cells(i.ColPrimaryLocation.Index).Value
    '                .Cells(17).Value = rowSumber.Cells(i.ColPrimaryUnit.Index).Value
    '                .Cells(18).Value = rowSumber.Cells(i.ColPrimarySecUnit.Index).Value
    '                .Cells(19).Value = -1
    '            End With
    '            'next row
    '            counter += 1
    '        End While

    '        If Obj.SimpanBaru1(vDGVStockOut, Me.CreatedBy) = True Then
    '            'Update primary stock out purchase return 
    '            Dim DTA1 As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
    '            Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn

    '            If DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary).HasValue = True Then
    '                Me.mPrimaryStockOutPurchaseReturn = DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
    '                DTA.UpdateQuery_PrimaryStockOutPurchaseReturn(Me.mPrimaryStockOutPurchaseReturn, Me.mPrimary)
    '                Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
    '            End If
    '            DTA1.Dispose()

    '            Obj.TransferJurnal(Obj.PrimarySupplier, Obj.Primary)

    '            Return True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        Obj = Nothing
    '        Frm.Dispose()
    '        ObjGRN = Nothing
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GenerateStockOutPurchaseReturn_QuickPR_Multi(ByVal vDGVItem As DataGridView, _
    ByVal i As FrmTransaksiPurchaseReturn_Quick, ByVal vNoTransaksiLama As String, _
    Optional ByVal vKoreksiProdQty As Boolean = True, Optional ByVal vSetPOB As Boolean = True, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim Obj As New Inventory_Module.clsTransaksiStockInOut
        Dim ObjGRN As New ClsTransaksiGRN
        Dim Frm As New Inventory_Module.FrmTransaksiStockOutPurchaseReturn2
        Dim vDGVStockOut As DataGridView = Frm.dgvItem
        Dim rowSumber, rowTarget As DataGridViewRow
        Dim counter As Double = 0
        Dim SecQty As Double
        Dim ED As Date
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim ObjSC As New Inventory_Module.ClsTransaksiStockCard

        Try
            Obj.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
            Obj.TanggalTransaksi = Me.TanggalTransaksi
            Obj.Remarks = ""
            Obj.SupplierDoNo = Me.SupplierDONumber
            Obj.PrimarySupplier = Me.mPrimarySupplier
            Obj.PrimaryPOPurchase = Me.mPrimary
            Obj.Primary = -1
            If vNoTransaksiLama <> "" Then
                Obj.NoTransaksi = vNoTransaksiLama
            Else
                Obj.NoTransaksi = Obj.GetNoBaru(Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn, _
                Me.TanggalTransaksi)
            End If

            If vFromCutOff = True Then
                Obj.CreatedBy = Me.mCreatedBy
                Obj.CreatedDate = Me.mCreatedDate
                Obj.ModifiedBy = Me.mModifiedBy
                Obj.ModifiedDate = Me.mModifiedDate
                Obj.TimeStamp = Me.mTimestamp
            End If

            'masukan data item ke dalam dgv
            vDGVStockOut.Rows.Clear()
            While counter < vDGVItem.Rows.Count
                rowSumber = vDGVItem.Rows(counter)
                vDGVStockOut.Rows.Add()
                rowTarget = vDGVStockOut.Rows(counter)
                With rowTarget
                    If IsNumeric(rowSumber.Cells(i.ColSecQty.Index).Value) = True Then
                        SecQty = rowSumber.Cells(i.ColSecQty.Index).Value
                    Else
                        SecQty = -1
                    End If
                    If IsDate(rowSumber.Cells(i.ColExpiryDate.Index).Value) = True Then
                        ED = rowSumber.Cells(i.ColExpiryDate.Index).Value
                    Else
                        ED = #1/1/1950#
                    End If

                    '.Cells(Me.ColNo.Index).Value = ""
                    '.Cells(Me.ColPONoDONo.Index).Value = Me.txtPurchaseOrderNo.Text
                    .Cells("ColSuppDONo").Value = Me.SupplierDONumber
                    '.Cells(Me.ColCodeNo.Index).Value = Me.TxtCodeNo.Text
                    '.Cells(Me.ColDescription.Index).Value = Me.txtDesc.Text
                    .Cells("ColQty").Value = rowSumber.Cells(i.ColTotalQty.Index).Value
                    '.Cells(Me.ColUnit.Index).Value = Me.lblUnitsDefault.Text
                    .Cells("ColSecondaryQty").Value = rowSumber.Cells(i.ColSecQty.Index).Value
                    '.Cells(Me.ColSecondaryUnit.Index).Value = Me.lblSecondaryUnits.Text
                    .Cells("ColBatch").Value = rowSumber.Cells(i.ColBatchNo.Index).Value
                    .Cells("ColExpiredDate").Value = ED
                    '.Cells(Me.ColLocation.Index).Value = Me.txtPrimaryLocation.Text
                    '.Cells(Me.ColWareHouse.Index).Value = Me.TxtWarehouse.Text
                    .Cells("ColRemark").Value = ""
                    .Cells("ColPrimaryPOnoDoNo").Value = rowSumber.Cells(i.ColPrimaryPO.Index).Value
                    .Cells("ColPrimaryCodeNo").Value = rowSumber.Cells(i.ColPrimaryCodeNo.Index).Value
                    .Cells("ColPrimaryLocation").Value = rowSumber.Cells(i.ColPrimaryLocation.Index).Value
                    .Cells("ColPrimaryUnits").Value = rowSumber.Cells(i.ColPrimaryUnit.Index).Value
                    .Cells("ColPrimarySecondaryUnits").Value = rowSumber.Cells(i.ColPrimarySecUnit.Index).Value
                    .Cells("ColPrimaryPdO").Value = -1

                    .Cells("ColQtyMulti").Value = rowSumber.Cells(i.ColTotalQtyMulti.Index).Value
                    '.Cells(Me.ColUnitCodeMulti.Index).Value = Me.lblUnitsMulti.Text
                    'If Me.txtPrimaryUnitsMulti.Text <> Me.txtPrimaryUnitsDefault.Text Then
                    '    .Cells(Me.ColRatioDesc.Index).Value = Me.ObjCD.DigitGrouping(Me.TxtRatioReference.Text, _
                    '            ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text, Me.txtPrimaryUnitsMulti.Text)) _
                    '            & " " & Me.lblUnitsMulti.Text & " = " & Me.ObjCD.DigitGrouping(Me.TxtRatioDefault.Text, _
                    '            Me.ObjProduct.GetAngkaDecimal(Me.TxtPrimaryCodeNo.Text)) & " " & Me.lblUnitsDefault.Text
                    'Else
                    '    .Cells(Me.ColRatioDesc.Index).Value = "-"
                    'End If
                    .Cells("ColRatioReference").Value = rowSumber.Cells(i.ColRatioReference.Index).Value
                    .Cells("ColRatioDefault").Value = rowSumber.Cells(i.ColRatioDefault.Index).Value
                    .Cells("ColPrimaryUnitMulti").Value = rowSumber.Cells(i.ColPrimaryUnitMulti.Index).Value

                End With
                'next row
                counter += 1
            End While
            Frm.TblItemCustom = i.TblItemInOutCustom.Copy
            If Obj.SimpanBaru1(vDGVStockOut, Me.CreatedBy, Frm.TblItemCustom, True, vSetPOB, -1, vCekDataBenar, vKoreksiProdQty, vFromCutOff) = True Then
                'Update primary stock out purchase return 
                Dim DTA1 As New DseTransaksiStockInOutTableAdapters.tbltransaksi_stockinoutTableAdapter
                Dim TipeTrans As Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn

                If DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary).HasValue = True Then
                    Me.mPrimaryStockOutPurchaseReturn = DTA1.ScalarQuery_GetPrimaryStockOutPurchaseReturn(TipeTrans, Me.mPrimary)
                    DTA.UpdateQuery_PrimaryStockOutPurchaseReturn(Me.mPrimaryStockOutPurchaseReturn, Me.mPrimary)
                    Me.mTimestamp = DTA.ScalarQuery_getTimeStamp(Me.mPrimary)
                    ObjSC.UpdateDocNo(Inventory_Module.ClsTransaksiStockCard.EnumJenisTransaksiAsal.StockOut_PurchaseReturn, Me.mPrimaryStockOutPurchaseReturn, Me.mNoTransaksi)
                End If
                DTA1.Dispose()

                'Obj.TransferJurnal(Obj.PrimarySupplier, Obj.Primary, -1)

                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Obj = Nothing
            Frm.Dispose()
            ObjGRN = Nothing
            DTA.Dispose()
            ObjSC = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        'tes aul
        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseReturn_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
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
                            Threading.Thread.Sleep(1000)
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

    Public Function Apakah_QuickPurchaseReturn(ByVal vPrimaryPR As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If DTA.ScalarQuery_GetQuickPurchase(vPrimaryPR).HasValue = True Then
                If DTA.ScalarQuery_GetQuickPurchase(vPrimaryPR).Value = 0 Then
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


    Private Sub CekNomor(ByVal DTA As DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter)
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
                        Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 5))
                    End If
                End If
                LoopCount += 1
            End While
            If LoopCount > 1 Then
                DTA.UpdateQuery_NoTrans_Urutan(Me.mNoTransaksi, Me.mNoUrut, Me.mPrimary)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCD = Nothing
        End Try
    End Sub

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseReturn_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
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
                            Threading.Thread.Sleep(1000)
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


    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, _
    ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, ByVal vTblCustom As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseReturn_SimpanPerubahan, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
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
                            Threading.Thread.Sleep(1000)
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

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, _
  ByVal vFrm As FrmTransaksiPurchaseReturn_Quick, _
  ByVal vTblCustom As CustomSMESS.DseTransaksiPurchaseReturnQuick_AddByCustomProduct.tbltrans_custom_PR_item1DataTableDataTable) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseReturn_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
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
                            Threading.Thread.Sleep(1000)
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

    Public Sub GetListPrimaryPO(ByVal vPrimaryPR As Double, ByVal vListPrimaryPO As List(Of Double))
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_item1TableAdapter
        Dim TBL As New DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item1DataTable
        Dim row As DseTransaksiPurchaseReturn.tbltrans_purchasereturn_item1Row

        Try
            vListPrimaryPO.Clear()

            DTA.FillBy_PrimaryPR(TBL, vPrimaryPR)
            For Each row In TBL.Rows
                vListPrimaryPO.Add(row.Primary_PO)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, ByVal vQuickPR As Boolean) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseReturn_Hapus, vUser)

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        Me.DTAInOUT.Delete(Me.vPrimary_Antrian)
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        Exit Function
                    Else
                        If Me.DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimaryTransaksi, vUser, vQuickPR)
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
                            Threading.Thread.Sleep(1000)
                        End If
                    End If
                End If
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, ByVal vUser As String, ByVal vQuickPR As Boolean) As Boolean
        Dim Obj As New ClsTransaksiPurchaseReturnItem
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim ObjItem As New ClsTransaksiPOItem
        Dim vListPrimaryPO As New List(Of Double)
        Dim TblItemCustom As New CustomSMESS.DseTransaksiStockInOut.tbltrans_custom_stockinout_itemDataTable

        Try
            If Me.ApakahBolehHapus(vUser) = True Then
                'hapus stock out purchase return, jika quick pr
                If vQuickPR = True Then
                    If Me.mPrimaryStockOutPurchaseReturn <> -1 Then
                        Dim ObjStockInOut As New Inventory_Module.clsTransaksiStockInOut
                        ObjStockInOut.CreatedBy = ObjStockInOut.GetCreatedBy(Me.mPrimaryStockOutPurchaseReturn)
                        ObjStockInOut.Primary = Me.mPrimaryStockOutPurchaseReturn
                        ObjStockInOut.TipeTransaksi = Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockOut_PurchaseReturn
                        If ObjStockInOut.Hapus1(ObjStockInOut.Primary, Me.mCreatedBy, TblItemCustom) = False Then
                            ObjStockInOut = Nothing
                            Return False
                        End If
                        Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                        ObjStockInOut = Nothing
                    End If
                End If

                Me.GetListPrimaryPO(vPrimaryTransaksi, vListPrimaryPO)
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    ObjItem.SetOrderBalance(vListPrimaryPO)
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Obj = Nothing
            ObjItem = Nothing
            vListPrimaryPO = Nothing
        End Try
    End Function

    Public Function GetPrimaryGRN(ByVal vPrimaryPR As Double) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturnTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryGRN(vPrimaryPR).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryGRN(vPrimaryPR).Value
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

    Public Function GetPrimary(ByVal vNoTransaksi As String) As Double
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getPrimary(vNoTransaksi)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_getPrimary(vNoTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksi(ByVal vPrimaryTransaksi As Double) As String
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_getNoTransaksi(vPrimaryTransaksi)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_getNoTransaksi(vPrimaryTransaksi)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal IsQuick As Integer, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.DataTable31TableAdapter

        Try
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                If DTA.getJmlBy_NoTransaksiLike(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.getJmlBy_NoTransaksiLike(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                If DTA.GetJmlBy_TglDariSampai(vTanggalDari, vTanggalSampai, IsQuick).HasValue = True Then
                    Return DTA.GetJmlBy_TglDariSampai(vTanggalDari, vTanggalSampai, IsQuick).Value
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                If DTA.getJmlBy_SupplierCodeLike(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.getJmlBy_SupplierCodeLike(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                If DTA.GetJmlBy_SupplierNameLike(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.GetJmlBy_SupplierNameLike(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoGRN Then
                If DTA.GetJmlBy_noGRNLike(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.GetJmlBy_noGRNLike(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoStockInPurchase Then
                If DTA.GetJmlBy_NoStockInPurchaseLike(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.GetJmlBy_NoStockInPurchaseLike(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                If DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", IsQuick).HasValue = True Then
                    Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%", IsQuick)
                Else
                    Return 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimary = -1
            mNoTransaksi = ""
            mNoUrut = -1
            mPrimaryStockOutPurchaseReturn = -1
            mPrimaryGRN = -1
            mTanggalTransaksi = Now
            mSubTotal = 0
            mTaxExclude = True
            mTax = 0
            mTaxAmount = 0
            mTotal = 0
            mReasonOfReturn = ""
            mRemarks = ""
            mSupplierDONumber = ""
            Me.mTipe_Return = EnumTipe_PurchaseReturn.Revisi_AOP
            Me.mTaxExchangeRate = 1
            Me.mDutystamp = 0
            Me.mPaymentTerm = 0
            Me.mPrimaryCurrency = -1

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    ''' <summary>
    ''' mengembalikan nilai true untuk po yang memiliki item yang menggunakan sizing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UseSizing(ByVal vPrimaryPR As Double) As Boolean
        Dim DTA As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn_itemTableAdapter

        Try
            If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryPR).HasValue = True Then
                If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryPR).Value > 0 Then
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


    ''' <summary>
    ''' mengambil item dari tabel dan mengumpulkannya berdasarkan category product yang terjadi
    ''' </summary>
    ''' <param name="vTBL"></param>
    ''' <remarks></remarks>
    Private Sub SetSizing(ByVal vTBL As DseTransaksiPurchaseReturn.TblCetakDataTable)
        Dim TBLDummy As New DseTransaksiPurchaseReturn.TblCetakDataTable
        Dim row, row1 As DseTransaksiPurchaseReturn.TblCetakRow
        Dim Filter As String = ""

        Try
            While vTBL.Rows.Count > 0
                row = vTBL.Rows(0)

                If row.Use_Sizing = 1 Then
                    Filter = "[Product Code] = '" + row.Category_Code + "' AND [Nett Price] = " + row.Nett_Price.ToString
                    TBLDummy.DefaultView.RowFilter = Filter
                    If TBLDummy.DefaultView.Count > 0 Then
                        'tambahkan qty
                        row1 = TBLDummy.DefaultView(0).Row
                        row1.Qty += row.Qty
                        'jika use sizing maka tambahkan ke sizing yang sesuai
                        If row.Use_Sizing = 1 Then
                            If row.Size = 36 Then
                                row1.Qty_36 += row.Total_Qty
                            ElseIf row.Size = 37 Then
                                row1.Qty_37 += row.Total_Qty
                            ElseIf row.Size = 38 Then
                                row1.Qty_38 += row.Total_Qty
                            ElseIf row.Size = 39 Then
                                row1.Qty_39 += row.Total_Qty
                            ElseIf row.Size = 40 Then
                                row1.Qty_40 += row.Total_Qty
                            ElseIf row.Size = 41 Then
                                row1.Qty_41 += row.Total_Qty
                            ElseIf row.Size = 42 Then
                                row1.Qty_42 += row.Total_Qty
                            ElseIf row.Size = 43 Then
                                row1.Qty_43 += row.Total_Qty
                            ElseIf row.Size = 44 Then
                                row1.Qty_44 += row.Total_Qty
                            ElseIf row.Size = 45 Then
                                row1.Qty_45 += row.Total_Qty
                            End If
                        End If

                        row1.Amount = row1.Nett_Price * row1.Qty
                        If row.Remarks <> "" Then
                            If row1.Remarks = "" Then
                                row1.Remarks = row.Remarks
                            Else
                                row1.Remarks = row1.Remarks + vbCrLf + row.Remarks
                            End If
                        End If
                    Else
                        'tambahkan bari baru
                        row1 = TBLDummy.NewTblCetakRow
                        'jika use sizing maka gunakan kode dan desc dari category
                        With row1
                            .Primary_Transaksi = row.Primary_Transaksi
                            .Amount = row.Amount
                            .Category_Code = row.Category_Code
                            .Category_Desc = row.Category_Desc
                            .Description = row.Category_Desc
                            .Discount = row.Discount
                            .Nett_Price = row.Nett_Price
                            .Product_Code = row.Category_Code
                            .Qty = row.Qty
                            .Qty_36 = 0
                            .Qty_37 = 0
                            .Qty_38 = 0
                            .Qty_39 = 0
                            .Qty_40 = 0
                            .Qty_41 = 0
                            .Qty_42 = 0
                            .Qty_43 = 0
                            .Qty_44 = 0
                            .Qty_45 = 0

                            .Remarks = row.Remarks
                            .Size = row.Size
                            .Unit_Code = row.Unit_Code
                            .Unit_Price = row.Unit_Price
                            .Use_Sizing = row.Use_Sizing
                            ._Decimal = row._Decimal
                            .Primary_PO = row.Primary_PO
                            .Total_Qty = row.Total_Qty

                            If row.Size = 36 Then
                                .Qty_36 = row.Total_Qty
                            ElseIf row.Size = 37 Then
                                .Qty_37 = row.Total_Qty
                            ElseIf row.Size = 38 Then
                                .Qty_38 = row.Total_Qty
                            ElseIf row.Size = 39 Then
                                .Qty_39 = row.Total_Qty
                            ElseIf row.Size = 40 Then
                                .Qty_40 = row.Total_Qty
                            ElseIf row.Size = 41 Then
                                row1.Qty_41 += row.Total_Qty
                            ElseIf row.Size = 42 Then
                                row1.Qty_42 += row.Total_Qty
                            ElseIf row.Size = 43 Then
                                row1.Qty_43 += row.Total_Qty
                            ElseIf row.Size = 44 Then
                                row1.Qty_44 += row.Total_Qty
                            ElseIf row.Size = 45 Then
                                row1.Qty_45 += row.Total_Qty
                            End If

                            .Amount = .Nett_Price * .Total_Qty
                        End With
                        TBLDummy.Rows.Add(row1)
                    End If
                Else
                    row = vTBL.Rows(0)
                    'tambahkan bari baru
                    row1 = TBLDummy.NewTblCetakRow
                    'jika use sizing maka gunakan kode dan desc dari category
                    With row1
                        .Primary_Transaksi = row.Primary_Transaksi
                        .Amount = row.Amount
                        .Category_Code = row.Category_Code
                        .Category_Desc = row.Category_Desc
                        .Description = row.Description
                        .Discount = row.Discount
                        .Nett_Price = row.Nett_Price
                        .Product_Code = row.Product_Code
                        .Qty = row.Qty
                        .Qty_36 = 0
                        .Qty_37 = 0
                        .Qty_38 = 0
                        .Qty_39 = 0
                        .Qty_40 = 0
                        .Qty_41 = 0
                        .Qty_42 = 0
                        .Qty_43 = 0
                        .Qty_44 = 0
                        .Qty_45 = 0
                        .Remarks = row.Remarks
                        .Size = row.Size
                        .Unit_Code = row.Unit_Code
                        .Unit_Price = row.Unit_Price
                        .Use_Sizing = row.Use_Sizing
                        ._Decimal = row._Decimal
                        .Primary_PO = row.Primary_PO
                        .Total_Qty = row.Total_Qty
                    End With
                    TBLDummy.Rows.Add(row1)
                End If

                vTBL.Rows.Remove(row)
            End While

            'pindahkan kembali
            While TBLDummy.Rows.Count > 0
                row = TBLDummy.Rows(0)
                'tambahkan bari baru
                row1 = vTBL.NewTblCetakRow
                'jika use sizing maka gunakan kode dan desc dari category
                With row1
                    .Primary_Transaksi = row.Primary_Transaksi
                    .Amount = row.Amount
                    .Category_Code = row.Category_Code
                    .Category_Desc = row.Category_Desc
                    .Description = row.Description
                    .Discount = row.Discount
                    .Nett_Price = row.Nett_Price
                    .Product_Code = row.Product_Code
                    .Qty = row.Qty
                    .Qty_36 = row.Qty_36
                    .Qty_37 = row.Qty_37
                    .Qty_38 = row.Qty_38
                    .Qty_39 = row.Qty_39
                    .Qty_40 = row.Qty_40
                    .Qty_41 = row.Qty_41
                    .Qty_42 = row.Qty_42
                    .Qty_43 = row.Qty_43
                    .Qty_44 = row.Qty_44
                    .Qty_45 = row.Qty_45
                    .Remarks = row.Remarks
                    .Size = row.Size
                    .Unit_Code = row.Unit_Code
                    .Unit_Price = row.Unit_Price
                    .Use_Sizing = row.Use_Sizing
                    ._Decimal = row._Decimal
                    .Primary_PO = row.Primary_PO
                    .Total_Qty = row.Total_Qty
                End With
                vTBL.Rows.Add(row1)
                TBLDummy.Rows.Remove(row)
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            TBLDummy.Dispose()
        End Try
    End Sub

    Private Function GetAddress(ByVal vTBL As DseTransaksiPurchaseReturn.TblGetIsiParameterCetakDataTable) As String
        Dim row As DseTransaksiPurchaseReturn.TblGetIsiParameterCetakRow
        Dim dummy As String = ""

        Try
            If vTBL.Rows.Count > 0 Then
                row = vTBL.Rows(0)
                If row.Address_1 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address_1
                    Else
                        dummy = dummy + vbCrLf + row.Address_1
                    End If
                End If
                If row.Address_2 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address_2
                    Else
                        dummy = dummy + vbCrLf + row.Address_2
                    End If
                End If
                If row.Address_3 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address_3
                    Else
                        dummy = dummy + vbCrLf + row.Address_3
                    End If
                End If
                If row.Address_4 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address_4
                    Else
                        dummy = dummy + vbCrLf + row.Address_4
                    End If
                End If
                If row.Address_5 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address_5
                    Else
                        dummy = dummy + vbCrLf + row.Address_5
                    End If
                End If

                Return dummy
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function Cetak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPO As Double, _
      ByVal vQuickPR As Boolean, ByVal vPrimarySupplier As Double) As Boolean
        Dim ObjSC As New Core_Module.ClsTransaksiSetting
        Dim ObjPR_Custom As New CustomSMESS.ClsTransPurchaseReturn
        If ObjSC.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PurchaseReturn, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            Return ObjPR_Custom.Cetak(CRV, vPrimaryPO, vQuickPR, vPrimarySupplier)
        End If

        If Me.UseSizing(vPrimaryPO) = True Then
            Return Me.Cetak2(CRV, vPrimaryPO, vQuickPR, vPrimarySupplier)
            Exit Function
        End If

        Dim DS As New DseTransaksiPurchaseReturn
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = ObjCur.GetDefaultPrecission(Me.GetPrimaryCurrency(vPrimaryPO))
        Dim row As DseTransaksiPurchaseReturn.TblGetIsiParameterCetakRow
        Dim Dummy As String = ""
        'Address Code
        Dim DTAGRN As New DseTransaksiPurchaseReturnTableAdapters.DataTable_AddressTableAdapter
        Dim TBLGRN As New DseTransaksiPurchaseReturn.DataTable_AddressDataTable
        Dim RowGRN As DseTransaksiPurchaseReturn.DataTable_AddressRow
        Dim Filter As String = ""
        Dim DTAPR As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter
        Dim CurrencyCode As String = ""
        Dim rowCetak As DseTransaksiPurchaseReturn.TblCetakRow
        Try
            Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.TblCetakTableAdapter
            DTA1.FillBy_PrimaryPO(DS.TblCetak, vPrimaryPO)
            'aul
            For Each rowCetak In DS.TblCetak
                With rowCetak
                    If .Primary_Units <> .Primary_Unit_Multi Then
                        .Decimal_Multi_ = ObjProduct.GetAngkaDecimal(.Primary_Product, .Primary_Unit_Multi)
                    Else
                        .Decimal_Multi_ = ._Decimal
                    End If

                End With
            Next
            '---
            DTA1.Dispose()
            Dim DTA2 As New DseTransaksiPurchaseReturnTableAdapters.TblGetIsiParameterCetakTableAdapter
            DTA2.Fill(DS.TblGetIsiParameterCetak, vPrimaryPO)
            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            'Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)
            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                row = DS.TblGetIsiParameterCetak.Rows(0)

                If vQuickPR = False Then
                    'Dapetin Address Dari GRN
                    DTAGRN.Fill(TBLGRN)
                    Filter = "[Primary] = " + row.Primary_GRN.ToString
                    TBLGRN.DefaultView.RowFilter = Filter
                    If TBLGRN.DefaultView.Count > 0 Then
                        RowGRN = TBLGRN.DefaultView(0).Row
                        row.Address_1 = RowGRN.Line_1
                        row.Address_2 = RowGRN.Line_2
                        row.Address_3 = RowGRN.Line_3
                        row.Address_4 = RowGRN.Line_4
                        row.Address_5 = RowGRN.Line_5
                        row.Supplier_Code = RowGRN.Supplier_Code
                        row.Supplier_Name = RowGRN.Supplier_Name
                    End If

                Else
                    'dapetin address dari supplier
                    Dim DTA_Supp As New DseTransaksiPurchaseReturnTableAdapters.DataTable_Address1TableAdapter
                    Dim TBL_Supplier As New DseTransaksiPurchaseReturn.DataTable_Address1DataTable
                    Dim row_Supp As DseTransaksiPurchaseReturn.DataTable_Address1Row
                    Dim ij As New FrmTransaksiPurchaseReturn_DetailAddress

                    DTA_Supp.Fill(TBL_Supplier, vPrimarySupplier)

                    If TBL_Supplier.Rows.Count = 0 Then
                        row.Address_1 = ""
                        row.Address_2 = ""
                        row.Address_3 = ""
                        row.Address_4 = ""
                        row.Address_5 = ""
                        row.Supplier_Code = ""
                        row.Supplier_Name = ""
                    Else
                        If TBL_Supplier.Rows.Count = 1 Then
                            row_Supp = TBL_Supplier.Rows(0)
                        Else
                            If String.IsNullOrEmpty(DTAPR.ScalarQuery_GetSupplierAddressIdentifier(vPrimaryPO)) = False Then
                                ij.SupplierAddressIdentifier = DTAPR.ScalarQuery_GetSupplierAddressIdentifier(vPrimaryPO)
                            End If
                            ij.TBL = TBL_Supplier
                            ij.ShowDialog()
                            row_Supp = TBL_Supplier.Rows(ij.cboAddressCode.SelectedIndex)
                        End If

                        'simpan data alamat yang digunakan pada cetakan
                        DTAPR.UpdateQuery_GetSupplierAddressIdentifier(row_Supp.Identifier, vPrimaryPO)

                        row.Address_1 = row_Supp.Line_1
                        row.Address_2 = row_Supp.Line_2
                        row.Address_3 = row_Supp.Line_3
                        row.Address_4 = row_Supp.Line_4
                        row.Address_5 = row_Supp.Line_5
                        row.Supplier_Code = row_Supp.Supplier_Code
                        row.Supplier_Name = row_Supp.Supplier_Name
                    End If

                    DTA_Supp.Dispose()
                    TBL_Supplier.Dispose()
                    ij = Nothing
                End If


                'Lanjutkan
                row = DS.TblGetIsiParameterCetak(0)
                If CurrencyCode = "" Then
                    CurrencyCode = row.Currency_Code
                End If
                If row.Address_1 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address_1
                    Else
                        Dummy = Dummy + vbCrLf + row.Address_1
                    End If
                End If
                If row.Address_2 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address_2
                    Else
                        Dummy = Dummy + vbCrLf + row.Address_2
                    End If
                End If
                If row.Address_3 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address_3
                    Else
                        Dummy = Dummy + vbCrLf + row.Address_3
                    End If
                End If
                If row.Address_4 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address_4
                    Else
                        Dummy = Dummy + vbCrLf + row.Address_4
                    End If
                End If
                If row.Address_5 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address_5
                    Else
                        Dummy = Dummy + vbCrLf + row.Address_5
                    End If
                End If


                Dim Report As New CrpCetakTransaksiPR

                Report.SetDataSource(DS)

                Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)
                Report.SetParameterValue(Report.Parameter_Currency.ParameterFieldName, ObjCur.GetCurrencyCode(Me.GetPrimaryCurrency(vPrimaryPO)))
                Report.SetParameterValue(Report.Parameter_CurrencyBase.ParameterFieldName, CurrencyCode)

                CRV.ReportSource = Report
                CRV.Show()
            End If
            DTA2.Dispose()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            'ObjPR_Custom = Nothing
            ObjSC = Nothing
            DTAPR = Nothing
        End Try
    End Function

    Private Function Cetak2(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPO As Double, _
    ByVal vQuickPR As Boolean, ByVal vPrimarySupplier As Double) As Boolean
        Dim DS As New DseTransaksiPurchaseReturn
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = ObjCur.GetDefaultPrecission(Me.GetPrimaryCurrency(vPrimaryPO))
        Dim row As DseTransaksiPurchaseReturn.TblGetIsiParameterCetakRow
        Dim DTA1 As New DseTransaksiPurchaseReturnTableAdapters.TblCetakTableAdapter
        Dim DTA2 As New DseTransaksiPurchaseReturnTableAdapters.TblGetIsiParameterCetakTableAdapter
        Dim Add As String = ""

        'Address Code
        Dim DTAGRN As New DseTransaksiPurchaseReturnTableAdapters.DataTable_AddressTableAdapter
        Dim TBLGRN As New DseTransaksiPurchaseReturn.DataTable_AddressDataTable
        Dim RowGRN As DseTransaksiPurchaseReturn.DataTable_AddressRow
        Dim Filter As String = ""
        Dim DTAPR As New DseTransaksiPurchaseReturnTableAdapters.tbltrans_purchasereturn1TableAdapter

        Try
            DTA1.FillBy_PrimaryPO(DS.TblCetak, vPrimaryPO)
            DTA1.Dispose()

            DTA2.Fill(DS.TblGetIsiParameterCetak, vPrimaryPO)
            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            'Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)

            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                'regroup berdasarkan sizing
                Me.SetSizing(DS.TblCetak)

                row = DS.TblGetIsiParameterCetak.Rows(0)
                'Dapetin Address Dari GRN
                If vQuickPR = False Then
                    'Dapetin Address Dari GRN
                    DTAGRN.Fill(TBLGRN)
                    Filter = "[Primary] = " + row.Primary_GRN.ToString
                    TBLGRN.DefaultView.RowFilter = Filter
                    If TBLGRN.DefaultView.Count > 0 Then
                        RowGRN = TBLGRN.DefaultView(0).Row
                        row.Address_1 = RowGRN.Line_1
                        row.Address_2 = RowGRN.Line_2
                        row.Address_3 = RowGRN.Line_3
                        row.Address_4 = RowGRN.Line_4
                        row.Address_5 = RowGRN.Line_5
                        row.Supplier_Code = RowGRN.Supplier_Code
                        row.Supplier_Name = RowGRN.Supplier_Name
                    End If
                Else
                    'dapetin address dari supplier
                    Dim DTA_Supp As New DseTransaksiPurchaseReturnTableAdapters.DataTable_Address1TableAdapter
                    Dim TBL_Supplier As New DseTransaksiPurchaseReturn.DataTable_Address1DataTable
                    Dim row_Supp As DseTransaksiPurchaseReturn.DataTable_Address1Row
                    Dim ij As New FrmTransaksiPurchaseReturn_DetailAddress

                    DTA_Supp.Fill(TBL_Supplier, vPrimarySupplier)

                    If TBL_Supplier.Rows.Count = 0 Then
                        row.Address_1 = ""
                        row.Address_2 = ""
                        row.Address_3 = ""
                        row.Address_4 = ""
                        row.Address_5 = ""
                        row.Supplier_Code = ""
                        row.Supplier_Name = ""
                    Else
                        If TBL_Supplier.Rows.Count = 1 Then
                            row_Supp = TBL_Supplier.Rows(0)
                        Else
                            If String.IsNullOrEmpty(DTAPR.ScalarQuery_GetSupplierAddressIdentifier(vPrimaryPO)) = False Then
                                ij.SupplierAddressIdentifier = DTAPR.ScalarQuery_GetSupplierAddressIdentifier(vPrimaryPO)
                            End If
                            ij.TBL = TBL_Supplier
                            ij.ShowDialog()
                            row_Supp = TBL_Supplier.Rows(ij.cboAddressCode.SelectedIndex)
                        End If

                        'simpan data alamat yang digunakan pada cetakan
                        DTAPR.UpdateQuery_GetSupplierAddressIdentifier(row_Supp.Identifier, vPrimaryPO)

                        row.Address_1 = row_Supp.Line_1
                        row.Address_2 = row_Supp.Line_2
                        row.Address_3 = row_Supp.Line_3
                        row.Address_4 = row_Supp.Line_4
                        row.Address_5 = row_Supp.Line_5
                        row.Supplier_Code = row_Supp.Supplier_Code
                        row.Supplier_Name = row_Supp.Supplier_Name
                    End If

                    DTA_Supp.Dispose()
                    TBL_Supplier.Dispose()
                    ij = Nothing
                End If

                row = DS.TblGetIsiParameterCetak(0)

                Dim Report As New CrpCetakTransaksiPR3

                Report.SetDataSource(DS)

                Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Me.GetAddress(DS.TblGetIsiParameterCetak))

                CRV.ReportSource = Report
                CRV.Show()
            End If
            DTA2.Dispose()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

   
End Class
