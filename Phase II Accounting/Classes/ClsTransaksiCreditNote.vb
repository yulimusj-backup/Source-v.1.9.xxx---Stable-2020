Option Explicit On
Option Strict Off
Public Class ClsTransaksiCreditNote
    Dim mPrimary As Double = -1
    Dim mPrimaryCustomer As Double = -1
    Dim mPrimaryAkunAr As Double = -1
    Dim mPrimaryExpenseAcc As Double = -1
    Dim mPrimaryCurrency As Double = -1
    Dim mPrimaryInvoice As Double = -1
    Dim mPrimaryJournal As Double = -1
    Dim mTgl As Date = Now
    Dim mUrutan As Double = -1
    Dim mNoTrans As String = ""
    Dim mIdentifierAddCust As String = ""
    Dim mCustomerDebitNote As String = ""
    Dim mAmount As Decimal = 0
    Dim mAmountReff As Decimal = 0
    Dim mExchangeRate As Decimal = 0
    Dim mDueDate As Date = Now
    Dim mReason As String = ""
    Dim mTipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE
    Dim mTipeKartu As ClsTransaksiCreditNote.EnumTipeKartu = EnumTipeKartu.Customer
    Dim mIsBeginingBalance As Integer
    Dim mContactPerson As String = ""
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim mNoJurnal As String = ""
    Dim mDescJurnal As String = ""
    Dim mReset As Integer = 0
    Dim mTglTransaksi As Date = Now
    Dim objCD As New Core_Module.ClsCekData
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering


    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Public Enum EnumTipeTransaksi
        CREDIT_NOTE
        DEBIT_NOTE
    End Enum
    Public Enum EnumTipeKartu
        Customer
        Supplier
    End Enum

    Public Property Primary() As Double
        Get
            Return mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return mPrimaryCustomer
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomer = value
        End Set
    End Property

    Public Property PrimaryExpenseAcc() As Double
        Get
            Return mPrimaryExpenseAcc
        End Get
        Set(ByVal value As Double)
            mPrimaryExpenseAcc = value
        End Set
    End Property

    Public Property PrimaryCurrency() As Double
        Get
            Return mPrimaryCurrency
        End Get
        Set(ByVal value As Double)
            mPrimaryCurrency = value
        End Set
    End Property

    Public Property PrimaryAkunAr() As Double
        Get
            Return mPrimaryAkunAr
        End Get
        Set(ByVal value As Double)
            mPrimaryAkunAr = value
        End Set
    End Property

    Public Property PrimaryInvoice() As Double
        Get
            Return mPrimaryInvoice
        End Get
        Set(ByVal value As Double)
            mPrimaryInvoice = value
        End Set
    End Property

    Public Property PrimaryJournal() As Double
        Get
            Return mPrimaryJournal
        End Get
        Set(ByVal value As Double)
            mPrimaryJournal = value
        End Set
    End Property

    Public Property Tgl() As Date
        Get
            Return mTgl
        End Get
        Set(ByVal value As Date)
            mTgl = value
        End Set
    End Property

    Public Property Urutan() As Double
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Double)
            mUrutan = value
        End Set
    End Property

    Public Property NoTrans() As String
        Get
            Return mNoTrans
        End Get
        Set(ByVal value As String)
            mNoTrans = value
        End Set
    End Property

    Public Property IdentifierAddCust() As String
        Get
            Return mIdentifierAddCust
        End Get
        Set(ByVal value As String)
            mIdentifierAddCust = value
        End Set
    End Property

    Public Property CustomerDebitNote() As String
        Get
            Return mCustomerDebitNote
        End Get
        Set(ByVal value As String)
            mCustomerDebitNote = value
        End Set
    End Property

    Public Property Amount() As Decimal
        Get
            Return mAmount
        End Get
        Set(ByVal value As Decimal)
            mAmount = value
        End Set
    End Property

    Public Property AmountReff() As Decimal
        Get
            Return mAmountReff
        End Get
        Set(ByVal value As Decimal)
            mAmountReff = value
        End Set
    End Property

    Public Property ExchangeRate() As Decimal
        Get
            Return mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
        End Set
    End Property

    Public Property DueDate() As Date
        Get
            Return mDueDate
        End Get
        Set(ByVal value As Date)
            mDueDate = value
        End Set
    End Property
    Public Property Reason() As String
        Get
            Return mReason
        End Get
        Set(ByVal value As String)
            mReason = value
        End Set
    End Property

    Public Property TipeTransaksi() As ClsTransaksiCreditNote.EnumTipeTransaksi
        Get
            Return mTipeTransaksi
        End Get
        Set(ByVal value As ClsTransaksiCreditNote.EnumTipeTransaksi)
            mTipeTransaksi = value
        End Set
    End Property

    Public Property TipeKartu() As ClsTransaksiCreditNote.EnumTipeKartu
        Get
            Return mTipeKartu
        End Get
        Set(ByVal value As ClsTransaksiCreditNote.EnumTipeKartu)
            mTipeKartu = value
        End Set
    End Property

    Public Property ContactPerson() As String
        Get
            Return mContactPerson
        End Get
        Set(ByVal value As String)
            mContactPerson = value
        End Set
    End Property
    Public Property IsBeginingBalance() As Integer
        Get
            Return mIsBeginingBalance
        End Get
        Set(ByVal value As Integer)
            mIsBeginingBalance = value
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
    Public Property CreatedDate() As Date
        Get
            Return mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
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
    Public Property NoJurnal() As String
        Get
            Return mNoJurnal
        End Get
        Set(ByVal value As String)
            mNoJurnal = value
        End Set
    End Property
    Public Property DescJurnal() As String
        Get
            Return mDescJurnal
        End Get
        Set(ByVal value As String)
            mDescJurnal = value
        End Set
    End Property
    Public Property Reset() As Integer
        Get
            Return mReset
        End Get
        Set(ByVal value As Integer)
            mReset = value
        End Set
    End Property
    Public Property TglTransaksi() As Date
        Get
            Return mTglTransaksi
        End Get
        Set(ByVal value As Date)
            mTglTransaksi = value
        End Set
    End Property
    Public Function GetTimeStamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTimestamp(CDec(vPrimary)).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
    Public Function GetPrimary(ByVal vNotrans As String) As Double
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNotrans).Value Then
                Return DTA.ScalarQuery_GetPrimary(vNotrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSuppInvoiceNoByPrimaryAOP(ByVal vPrimaryAop As Double) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Not String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierInvoiceNo(vPrimaryAop)) Then
                Return DTA.ScalarQuery_GetSupplierInvoiceNo(vPrimaryAop).ToString
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


    Public Function GetSuppInvoiceNoByPrimaryCDNote(ByVal vPrimaryCDNote As Double) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim PrimaryAOP As Double = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Not String.IsNullOrEmpty(DTA.ScalarQuery_GetSupplierInvoiceNo(vPrimaryCDNote)) Then
                Return DTA.ScalarQuery_GetSupplierInvoiceNo(vPrimaryCDNote)
            End If
            'If PrimaryAOP <> -1 Then
            '    Return Me.GetSuppInvoiceNoByPrimaryAOP(PrimaryAOP)
            'End If

            Return ""

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAOPInvoice(ByVal vPrimary As Double, ByVal vtipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi) As Decimal
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlAOPInvoice(vPrimary, vtipeTransaksi).HasValue Then
                Return DTA.ScalarQuery_GetJmlAOPInvoice(vPrimary, vtipeTransaksi).Value
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTrans(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNotransByPrimary(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub GetNoTransaksi_Remarks_Tanggal(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByRef vDate As Date, ByVal vPrimary As Double, ByRef vDueDate As Date)
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL As New DseTransaksiCreditNote.tbltrans_credit_noteDataTable
        Dim row As DseTransaksiCreditNote.tbltrans_credit_noteRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.reason
                vNoTransaksi = row.no_transksi
                vDate = row.tgl
                vDueDate = row.due_date
            Else
                vRemarks = ""
                vNoTransaksi = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub


    Public Sub GetNoTransaksi_Remarks_Tanggal(ByRef vNoTransaksi As String, ByRef vRemarks As String, ByRef vDate As Date, ByVal vPrimary As Double, ByRef vSupplierInvoiceNo As String, ByRef vDueDate As Date)
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TBL As New DseTransaksiCreditNote.tbltrans_credit_noteDataTable
        Dim row As DseTransaksiCreditNote.tbltrans_credit_noteRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)
            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                vRemarks = row.reason
                vNoTransaksi = row.no_transksi
                vDate = row.tgl
                vSupplierInvoiceNo = row.customer_debet_note
                vDueDate = row.due_date
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    'Public Function GetAccountCode(ByVal vPrimary As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Dim Retval As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        Retval = DTA.ScalarQuery_GetArAccountCode(vPrimary)
    '        If String.IsNullOrEmpty(Retval) = True Then
    '            Return ""
    '        Else
    '            Return Retval
    '        End If
    '        Return ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    'Public Function GetCurrCode(ByVal vPrimary As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Dim Retval As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        Retval = DTA.ScalarQuery_GetCurrCode(vPrimary)
    '        If String.IsNullOrEmpty(Retval) = True Then
    '            Return ""
    '        Else
    '            Return Retval
    '        End If
    '        Return ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetInvoiceNumber(ByVal vPrimary As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Dim Retval As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        Retval = DTA.ScalarQuery_GetInvoiceNumber(vPrimary)
    '        If String.IsNullOrEmpty(Retval) = True Then
    '            Return ""
    '        Else
    '            Return Retval
    '        End If
    '        Return ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    Public Function GetJml(ByVal vUnitCode As String) As Double
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetJml(vUnitCode)) = False Then
                Return DTA.ScalarQuery_GetJml(vUnitCode)
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

    Public Function Get_Jml(ByRef vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria, _
                                 Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
                                 Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.DataTable1TableAdapter
        Dim RetVal As Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vSearchBy = EnumSearchCriteria.ByNoTransaksi Then
                RetVal = DTA.ScalarQuery_GetJmlByNoTransLIKE(vKriteria + "%", TipeTransaksi)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDate Then
                RetVal = DTA.ScalarQuery_GetJmlByTgldarisampai(TipeTransaksi, vTanggalDari, vTanggalSampai)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                RetVal = DTA.ScalarQuery_GetJmlByReasonLike(TipeTransaksi, vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByAccountNo Then
                RetVal = DTA.ScalarQuery_GetJmlByNoJurnalLIKE(TipeTransaksi, vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            ElseIf vSearchBy = EnumSearchCriteria.ByDescJurnal Then
                RetVal = DTA.ScalarQuery_GetJmlByDescJurnal(TipeTransaksi, 0, vKriteria + "%")
                If RetVal.HasValue = True Then
                    Return RetVal.Value
                Else
                    Return 0
                End If
            Else
                RetVal = DTA.ScalarQuery_GetJmlByRefLIKE(vKriteria + "%", TipeTransaksi)
                If RetVal.HasValue = True Then
                    Return RetVal.Value
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

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vTipeTransaksi As EnumTipeTransaksi) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note, _
                vTanggalTransaksi)
            Else
                Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note, _
                vTanggalTransaksi)
            End If

            'cari no urut
            If DTA.ScalarQuery_GetMaxUrutan(CInt(vTipeTransaksi), TanggalStart, TanggalEnd).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan(CInt(vTipeTransaksi), TanggalStart, TanggalEnd).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note)
            Else
                Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetNoBaru(ByVal vTanggalTransaksi As Date, ByVal vTipeNoTrans As String, ByVal vTipeTransaksi As Integer) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    'Dim TipeTransaksi As Integer
    '    Dim TanggalStart As Date
    '    Dim TanggalEnd As Date
    '    Dim NoUrut As String
    '    Dim TipeStr As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        'cari tanggal sampai dan dari
    '        TanggalEnd = objCD.GetLastDateOfTheYear(vTanggalTransaksi)
    '        TanggalStart = objCD.GetFisrtDateOfTheYear(vTanggalTransaksi)
    '        'cari no urut
    '        If DTA.ScalarQuery_GetMaxUrutan(vTipeTransaksi, TanggalStart, TanggalEnd).HasValue = True Then
    '            NoUrut = CStr(DTA.ScalarQuery_GetMaxUrutan((vTipeTransaksi), TanggalStart, TanggalEnd).Value + 1)
    '            If NoUrut <= 0 Then
    '                NoUrut = 1
    '            End If
    '        Else
    '            NoUrut = 1
    '        End If
    '        While NoUrut.Length < 6
    '            NoUrut = "0" + NoUrut
    '        End While
    '        Return NoUrut + "/" + vTipeNoTrans + "/" + CStr(Format(vTanggalTransaksi, "MM")) + "/" + _
    '        CStr(Format(vTanggalTransaksi, "yyyy"))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetPrimaryakunByNoInvoice(ByVal vNoInvoice As String) As Double
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryAkunArByInvoiceNumber(vNoInvoice).HasValue Then
    '            Return DTA.ScalarQuery_GetPrimaryAkunArByInvoiceNumber(vNoInvoice).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function GetPrimarCurrencyByKodeAkunAr(ByVal vAccountCode As String) As Double
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblisi_cbo_ar_akunTableAdapter
    '    Dim PrimaryCurrencyDummy As Nullable(Of Long)
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        PrimaryCurrencyDummy = DTA.ScalarQuery_GetPrimaryCurrencyByAkunCode(vAccountCode)
    '        If PrimaryCurrencyDummy.HasValue Then
    '            Return PrimaryCurrencyDummy.Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return -1
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    'Public Function Get_DefaultExchangerateByInvoice(ByVal vPrimary As Double, ByVal vtipetransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Dim Retval As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If vtipetransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
    '            Retval = DTA.ScalarQuery_GetTaxExchangeratebyPrimaryInvoice(CDec(vPrimary))
    '        Else
    '            Retval = DTA.ScalarQuery_GetTaxExchangeratebyPrimaryInvoice(CDec(vPrimary))
    '        End If

    '        If String.IsNullOrEmpty(Retval) = True Then
    '            Return ""
    '        Else
    '            Return Retval
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function Get_TotalAmountByAOP(ByVal vPrimary As Double, ByVal vtipetransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vtipetransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Retval = DTA.ScalarQuery_GetTotalByPrimaryAOP(CDec(vPrimary))
            Else
                Retval = DTA.ScalarQuery_GetTotalByPrimaryAOP(CDec(vPrimary))
            End If

            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_TotalBySales(ByVal vPrimary As Double, ByVal vtipetransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vtipetransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Retval = DTA.ScalarQuery_GetTotalByPrimarySales(CDec(vPrimary))
            Else
                Retval = DTA.ScalarQuery_GetTotalByPrimarySales(CDec(vPrimary))
            End If

            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Get_DefaultExchangerateByPrimaryCurrency(ByVal vPrimaryCurrency As Double) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblmaster_currencyTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimaryCurrency)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetDescExpenseAkun(ByVal vPrimary As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblisi_cbo_ar_akunTableAdapter
    '    Try
    '        If String.IsNullOrEmpty(DTA.ScalarQuery_GetDescbyPrimaryAkun(vPrimary)) = False Then
    '            Return DTA.ScalarQuery_GetDescbyPrimaryAkun(vPrimary)
    '        Else
    '            Return ""
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    'Public Function GetAccountCodeByPrimaryAkunar(ByVal vPrimary As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblisi_cbo_ar_akunTableAdapter
    '    Dim Retval As String = ""
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        Retval = DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)
    '        If String.IsNullOrEmpty(Retval) = True Then
    '            Return ""
    '        Else
    '            Return Retval
    '        End If
    '        Return ""
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    'Public Function GetPrimaryCustomerByNoInvoice(ByVal vNoInvoice As String) As Double
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryCustomerByNoInvoice(vNoInvoice).Value Then
    '            Return DTA.ScalarQuery_GetPrimaryCustomerByNoInvoice(vNoInvoice).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    'Public Function GetPrimaryByAccountCode(ByVal vAccountCode As String) As Double
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblmaster_akunTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryByAccountCode(vAccountCode).Value Then
    '            Return DTA.ScalarQuery_GetPrimaryByAccountCode(vAccountCode).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
    'Public Function GetPrimaryByNoAOP(ByVal vNoAop As String) As Double
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring
    '        If DTA.ScalarQuery_GetPrimaryByNoAOP(vNoAop).Value Then
    '            Return DTA.ScalarQuery_GetPrimaryByNoAOP(vNoAop).Value
    '        Else
    '            Return -1
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function

    Public Function GetPrimarySupplierAOP(ByVal vPrimaryAOP As Double) As Double
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_aopTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimarySupplierBy(vPrimaryAOP).HasValue AndAlso _
               DTA.ScalarQuery_GetPrimarySupplierBy(vPrimaryAOP).Value > 0 Then
                Return DTA.ScalarQuery_GetPrimarySupplierBy(vPrimaryAOP).Value
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

    Public Function Get_DefaultExchangerateByPrimary(ByVal vPrimary As Double, ByVal vtipetransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblmaster_currencyTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring



            If vtipetransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                If DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).HasValue AndAlso _
                   DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).Value > 0 Then
                    Retval = DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).Value
                Else
                    Retval = ""
                End If
            Else
                If DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).HasValue AndAlso _
                   DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).Value > 0 Then
                    Retval = DTA.ScalarQuery_GetDefaultExchangerateByPrimaryCurrency(vPrimary).Value
                Else
                    Retval = ""
                End If
            End If



            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCustInvoice(ByVal vPrimaryInv As Double) As Double
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_salesinvoiceTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCust(vPrimaryInv).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCust(vPrimaryInv).Value
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

    Public Function GetPrimaryCurrencyAPAR(ByVal vPrimaryTransaksi As Double) As Double
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCurrencyAR(vPrimaryTransaksi).Value Then
                Return DTA.ScalarQuery_GetPrimaryCurrencyAR(vPrimaryTransaksi).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'Public Function GetPrimaryCurrency(ByVal vPrimaryCDNote As Double) As String
    '    Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
    '    Dim Primary As Double = -1
    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring

    '        If DTA.ScalarQuery_GetTaxExchangeratebyPrimaryInvoice(vPrimaryCDNote).HasValue AndAlso _
    '           DTA.ScalarQuery_GetTaxExchangeratebyPrimaryInvoice(vPrimaryCDNote).Value > -1 Then
    '            Primary = DTA.ScalarQuery_GetTaxExchangeratebyPrimaryInvoice(vPrimaryCDNote).Value
    '        End If

    '        Return Primary

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '        Return ""
    '    Finally
    '        DTA.Dispose()
    '    End Try
    'End Function
#Region "Isi Tbl"
    Public Sub IsiTBLArAcount(ByVal vPerkiraan1 As ClsMasterAkun.EnumTipeAkun, ByVal vPerkiraan2 As ClsMasterAkun.EnumTipeAkun2, _
       ByVal TBL As DseTransaksiCreditNote.tblisi_cbo_ar_akunDataTable)
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblisi_cbo_ar_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill_ByTipePerkiraan(TBL, vPerkiraan1, vPerkiraan2)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLAOPAcount(ByVal vPrimary As String, _
    ByVal TBL As DseTransaksiCreditNote.tblmaster_akunDataTable)
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiTBLCurrency(ByVal vPrimary As Double, ByVal TBL As DseTransaksiCreditNote.tblmaster_currencyDataTable)
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tblmaster_currencyTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub
#End Region
    Overridable Function SimpanBaru(ByVal vUserCode As String, ByVal vFrm As FrmTransaksiCreditNote) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)

            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.Primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.CreditDebitNote_SimpanBaru, vUserCode)
            'Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                If SW.Elapsed.TotalSeconds >= ObjAntrin.GetWaitTime Then
                    MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                    DTAInOUT.Delete(Me.vPrimary_Antrian)
                    Exit Function
                Else
                    'hanya mencoba ke database setelah interval waktu tertentu saja
                    If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanBaru1(vUserCode, vFrm)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Private Function ApakahDataBenar(ByVal vDate As Date, ByVal vPrimaryCurrency As Double, ByVal vBaru As Boolean, ByVal vTipeTransaksi As EnumTipeTransaksi) As Boolean
        Dim ObjJournal As New ClsTransaksiJournal
        Dim TipeTransaksi1 As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note

        Try
            'tidak boleh ada journal revaluasi untuk tgl ybs
            If vPrimaryCurrency <> Core_Module.PrimaryBaseCurrency Then
                If ObjJournal.ApakahSudahRevaluasi(vDate) = True Then
                    MsgBox("There's end of month journal for period " + Format(vDate, "MMM yyyy") + _
                    ", you can't create/edit credit note/debit note with non-IDR currency.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'terkait penomoran
            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                TipeTransaksi1 = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeTransaksi1 = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If

            If vBaru = True Then
                If Me.ObjNumbering.Apakah_Autonumber(TipeTransaksi1) = True Then
                    If Me.ApakahSudahAda(Me.mNoTrans) = True Then
                        Me.mNoTrans = Me.GetNoBaru(Me.mTgl, vTipeTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTrans) = True Then
                        MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(TipeTransaksi1) = True Then
                    If Me.ApakahSudahAda(Me.mNoTrans) = True Then
                        If Me.GetPrimary(Me.mNoTrans) <> Me.mPrimary Then
                            Me.mNoTrans = Me.GetNoBaru(Me.mTgl, vTipeTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTrans) = True Then
                        If Me.GetPrimary(Me.mNoTrans) <> Me.mPrimary Then
                            MsgBox("Document number duplication, please change document number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjJournal = Nothing
        End Try
    End Function

    Overridable Function SimpanBaru1(ByVal vUserCode As String, ByVal vFrm As FrmTransaksiCreditNote, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA1 As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_note2TableAdapter
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim Timestamp_ As Date = Now
        Dim Urutan As Double = 0
        Dim Reason1 As String = ""
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Decimal = 0
        Try
            DefPrec = ObjCurr.GetDecimal_DefaultCurrency
            If vCekDataBenar = True Then
                If Me.ApakahDataBenar(Me.mTgl, Me.mPrimaryCurrency, True, Me.mTipeTransaksi) = False Then
                    Return False
                End If
            End If

            If vFromCutOff = True Then
                CreatedBy_ = Me.mCreatedBy
                CreatedDate_ = Me.mCreatedDate
                ModifiedBy_ = Me.mModifiedBy
                ModifiedDate_ = Me.mModifiedDate
                Timestamp_ = Me.mTimestamp
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '- isikan urutan nilai index pertama dan batas panjang nilainya
            If Me.mTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Me.mUrutan = Me.ObjNumbering.GetUrutan(Me.mNoTrans, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note, _
                    Me.mTgl)
            Else
                Me.mUrutan = Me.ObjNumbering.GetUrutan(Me.mNoTrans, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note, _
                    Me.mTgl)
            End If

            If DTA1.InsertQuery(Me.mNoTrans, Me.mTgl, Me.mUrutan, Me.mPrimaryCustomer, _
                        Me.mPrimaryAkunAr, Me.mIdentifierAddCust, Me.mCustomerDebitNote, Me.mPrimaryInvoice, _
                        Me.mPrimaryExpenseAcc, Me.mPrimaryCurrency, Me.mAmount, Me.mDueDate, _
                        Me.mReason, Me.mTipeTransaksi, Me.mTipeKartu, CreatedBy_, _
                        CreatedDate_, ModifiedBy_, ModifiedDate_, Timestamp_, Me.mIsBeginingBalance, _
                        Me.mContactPerson, Me.mExchangeRate, Me.mPrimaryJournal, Me.mAmountReff, Me.mReset, Me.mNoJurnal, Me.mDescJurnal, Me.mTglTransaksi) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mNoTrans)
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                Dim TipeJurnal2 As New ClsTransaksiJournal.EnumTipeJurnal2
                Dim JenisTransaksiAsal As New ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
                If Me.mTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Credit_Note
                    JenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
                Else
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Debit_Note
                    JenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
                End If

                Dim NilaiRupiah As Decimal = 0
                NilaiRupiah = Me.objCD.DigitGrouping(CDec(Me.mAmount * Me.mExchangeRate), DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                If Me.mTipeKartu = EnumTipeKartu.Customer Then
                    Dim primaryJournal As Double = -1
                    If Me.CustomerDebitNote.Trim <> "" Then
                        Reason1 = Me.mReason + "(" + Me.CustomerDebitNote.Trim + ")"
                    End If
                    primaryJournal = Me.TransferJurnalCustomer(Me.mPrimary, Me.mPrimaryCustomer, _
                                                                     NilaiRupiah, Me.mAmountReff, 0, _
                                                                     Me.mUrutan, Me.mTgl, _
                                                                     Me.NoTrans, vUserCode, _
                                                                     (Me.mDueDate.Date - Me.mTgl.Date).Days, Me.mExchangeRate, _
                                                                     Me.mAmount, Me.PrimaryCurrency, _
                                                                     Me.PrimaryExpenseAcc, Me.mPrimaryJournal, Me.mPrimaryAkunAr, TipeJurnal2, _
                                                                     JenisTransaksiAsal, Me.mTipeTransaksi, Reason1, vFrm)
                    If primaryJournal > -1 Then
                        Me.mPrimaryJournal = primaryJournal
                        DTA.UpdateQuery_PrimaryJournal(Me.mPrimaryJournal, Me.mPrimary)

                    End If

                ElseIf Me.mTipeKartu = EnumTipeKartu.Supplier Then
                    Dim primaryJournal As Double = -1
                    If Me.CustomerDebitNote.Trim <> "" Then
                        Reason1 = Me.mReason + "(" + Me.CustomerDebitNote.Trim + ")"
                    End If
                    primaryJournal = Me.TransferJurnalSupplier(Me.mPrimary, TipeJurnal2, _
                                              JenisTransaksiAsal, Me.mTipeTransaksi, _
                                              Me.mPrimaryCustomer, NilaiRupiah, Me.mAmountReff, 0, _
                                              Me.mUrutan, Me.mTgl, _
                                              Me.NoTrans, vUserCode, _
                                               (Me.mDueDate.Date - Me.mTgl.Date).Days, Me.mAmount, _
                                                Me.mExchangeRate, Me.PrimaryCurrency, _
                                               Me.mPrimaryInvoice, Me.mPrimaryAkunAr, Me.mPrimaryExpenseAcc, Reason1, vFrm)
                    If primaryJournal > -1 Then
                        Me.mPrimaryJournal = primaryJournal
                        DTA.UpdateQuery_PrimaryJournal(Me.mPrimaryJournal, Me.mPrimary)
                    End If
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTrans + "' for key 3 " Then
                Return Me.SimpanBaru(vUserCode, vFrm)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String, ByVal vFrm As FrmTransaksiCreditNote) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)

            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.CreditDebitNote_SimpanBaru, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.SimpanPerubahan1(vUserCode, vFrm)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function SimpanPerubahan1(ByVal vUserCode As String, ByVal vFrm As FrmTransaksiCreditNote) As Boolean
        Dim DTA1 As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_note2TableAdapter
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Reason1 As String = ""
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Decimal = 0
        Try
            DefPrec = ObjCurr.GetDecimal_DefaultCurrency
            If Me.ApakahDataBenar(Me.mTgl, Me.mPrimaryCurrency, False, Me.mTipeTransaksi) = False Then
                Return False
            End If

            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.mTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Me.mUrutan = Me.ObjNumbering.GetUrutan(Me.mNoTrans, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note, _
                    Me.mTgl)
            Else
                Me.mUrutan = Me.ObjNumbering.GetUrutan(Me.mNoTrans, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note, _
                    Me.mTgl)
            End If
            If DTA1.UpdateQuery(Me.mNoTrans, mTgl, mUrutan, _
                          Me.mPrimaryCustomer, Me.mPrimaryAkunAr, Me.mIdentifierAddCust, _
                          Me.mCustomerDebitNote, Me.mPrimaryInvoice, Me.mPrimaryExpenseAcc, _
                          Me.mPrimaryCurrency, Me.mAmount, Me.mDueDate, _
                          Me.mReason, Me.mTipeTransaksi, Me.mTipeKartu, _
                          vUserCode, Now, Now, Me.mIsBeginingBalance, _
                          Me.ContactPerson, Me.mExchangeRate, Me.mPrimaryJournal, Me.mAmountReff, Me.mReset, Me.mNoJurnal, Me.mDescJurnal, Me.mTglTransaksi, Me.mPrimary) > 0 Then
                Me.mTimestamp = Me.GetTimeStamp(Me.mPrimary) 'DTA.ScalarQuery_GetTimestamp(CDec(Me.mPrimary)).Value
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode

                Dim TipeJurnal2 As New ClsTransaksiJournal.EnumTipeJurnal2
                Dim JenisTransaksiAsal As New ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal
                If Me.mTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Credit_Note
                    JenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note
                Else
                    TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Debit_Note
                    JenisTransaksiAsal = ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note
                End If


                Dim NilaiRupiah As Decimal = 0
                NilaiRupiah = Me.objCD.DigitGrouping(CDec(Me.mAmount * Me.mExchangeRate), DefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                If Me.mTipeKartu = EnumTipeKartu.Customer Then
                    Dim primaryJournal As Double = -1
                    If Me.CustomerDebitNote.Trim <> "" Then
                        Reason1 = Me.mReason + "(" + Me.CustomerDebitNote.Trim + ")"
                    End If
                    primaryJournal = Me.TransferJurnalCustomer(Me.mPrimary, Me.mPrimaryCustomer, _
                                                                     NilaiRupiah, Me.mAmountReff, 0, _
                                                                     Me.mUrutan, Me.mTgl, _
                                                                     Me.NoTrans, vUserCode, _
                                                                     (Me.mDueDate.Date - Me.mTgl.Date).Days, Me.mExchangeRate, _
                                                                     Me.mAmount, Me.PrimaryCurrency, _
                                                                     Me.PrimaryExpenseAcc, Me.mPrimaryJournal, Me.mPrimaryAkunAr, TipeJurnal2, _
                                                                     JenisTransaksiAsal, Me.mTipeTransaksi, Reason1, vFrm)
                    If primaryJournal > -1 Then
                        Me.mPrimaryJournal = primaryJournal
                        DTA.UpdateQuery_PrimaryJournal(Me.mPrimaryJournal, Me.mPrimary)
                    End If

                ElseIf Me.mTipeKartu = EnumTipeKartu.Supplier Then
                    Dim primaryJournal As Double = -1
                    If Me.CustomerDebitNote.Trim <> "" Then
                        Reason1 = Me.mReason + "(" + Me.CustomerDebitNote.Trim + ")"
                    End If
                    primaryJournal = Me.TransferJurnalSupplier(Me.mPrimary, TipeJurnal2, _
                                              JenisTransaksiAsal, Me.mTipeTransaksi, _
                                              Me.mPrimaryCustomer, NilaiRupiah, Me.mAmountReff, 0, _
                                              Me.mUrutan, Me.mTgl, _
                                              Me.NoTrans, vUserCode, _
                                               (Me.mDueDate.Date - Me.mTgl.Date).Days, Me.mAmount, _
                                                Me.mExchangeRate, Me.PrimaryCurrency, _
                                               Me.mPrimaryInvoice, Me.mPrimaryAkunAr, Me.mPrimaryExpenseAcc, Reason1, vFrm)
                    If primaryJournal > -1 Then
                        Me.mPrimaryJournal = primaryJournal
                        DTA.UpdateQuery_PrimaryJournal(Me.mPrimaryJournal, Me.mPrimary)
                    End If
                End If
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

#Region "Transfer Journal"
    Private Function TransferJurnalSupplier(ByVal vPrimaryCDNote As Double, _
      ByVal vTipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2, _
      ByVal vTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
      ByVal vTipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi, _
      ByVal vPrimarySupp As Double, ByVal vNilaiCDNote As Decimal, ByVal vNilaiCDNoteReff As Decimal, _
      ByVal vVAT As Decimal, ByVal vUrutanCDNote As Double, _
      ByVal vTanggalCDNote As Date, ByVal vCDNoteNumber As String, _
      ByVal vUserCode As String, ByVal vPaymentTerm As Double, _
      ByVal vCDNoteOriginalAmount As Decimal, ByVal vExchangeRate As Decimal, _
      ByVal vPrimaryCurrency As Double, ByVal vPrimaryAOP As Double, ByVal vPrimaryAkunAP As Double, _
      ByVal vPrimaryAkunReff As Double, ByVal strRemarks As String, ByVal vFrm As FrmTransaksiCreditNote) As Double


        Dim DTA_Balance As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim TipeBalance As ClsTransaksiAOP.EnumTipePurchaseBalance
        Dim DTA_Jurnal As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim RefNumber As String = CStr(vUrutanCDNote)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit, DebetOriginal, CreditOriginal As Decimal
        Dim rowDGV As DataGridViewRow
        Dim DTAAkun As New DseTransaksiCreditNote_TransferJournalTableAdapters.tblmaster_akunTableAdapter
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim Local As Boolean = False
        Dim LocalStr As String = ""
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim SuppCode, SuppName As String
        Dim MonthStr As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim AP_PadaCost As Decimal = 0
        Dim AP_PadaCost_Original As Decimal = 0
        Dim AP_PadaCost_Base As Decimal = 0
        Dim PrimaryAkunAP As Double = -1
        Dim PrimaryAkunAP_BaseCurrency As Double = -1
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkun As Double = -1
        'Dim DTAInv As New DseTransaksiCreditNote_TransferJournalTableAdapters.TBLInventoryTableAdapter
        'Dim TBLInv As New DseTransaksiCreditNote_TransferJournal.TBLInventoryDataTable
        'Dim rowInv As DseTransaksiCreditNote_TransferJournal.TBLInventoryRow
        Dim ListAkunInv As New List(Of Double)
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrecBase As Integer = 0
        Dim SelisihDP As Decimal = 0
        Dim Dummy1 As Decimal = 0
        Dim ListPrimaryAkunCost As New List(Of Double)
        Dim DefPrecCur As Integer = 0
        Dim IntegerDummy, TotalDummy, SubTotalDummy As Integer
        Dim Cost As Decimal = 0
        Dim ObjAkun As New ClsMasterAkun
        Dim NewPrimaryJournal As Double = -1
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Prefix As String = ""
        Dim Counter As Integer = 0
        Dim RefNumber1 As String = ""
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
        Dim vRemarks As String = ""
        Dim vDgv As DataGridView = vFrm.dgvItem
        Dim vRowDgv As New DataGridViewRow
        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'DTAInv.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Jurnal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Balance.Connection.ConnectionString = DTA_Jurnal.Connection.ConnectionString
            DTAAkun.Connection.ConnectionString = DTA_Jurnal.Connection.ConnectionString

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, vTipeJurnal2, vPrimaryCDNote)
            'hapuskan data balance sales lama
            'If vPrimaryAOP > -1 Then
            '    DTA_Balance.Delete(vPrimaryCDNote, TipeBalance)
            'Elses
            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If
            DTA_Balance.Delete(vPrimaryCDNote, TipeBalance)
            'End If


            '---------------------------------------------------------------------------------'
            '-- hapus data kartu AP dan AR ---------------------------------------------------'
            '                                                                                 '
            ObjKartu.Hapus(vPrimaryCDNote, vTransaksiAsal)
            '                                                                                 '
            '---------------------------------------------------------------------------------'

            If Me.ApakahTransferOtomatis(vTanggalCDNote) = True Then
                SuppCode = ObjSupplier.GetSupplierCode(vPrimarySupp)
                SuppName = ObjSupplier.GetSupplierName(vPrimarySupp)
                '---------------------------------------------------------------------------------'
                'buat jurnal baru
                '                                                                                 '
                Dim strDesc As String = ""
                'siapkan penomoran jurnal yang terkait
                'If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
                '    If RefNumber.Length < 6 Then
                '        While RefNumber.Length < 6
                '            RefNumber = "0" + RefNumber
                '        End While
                '    End If
                '    MonthStr = vTanggalCDNote.Month.ToString
                '    While MonthStr.Length < 2
                '        MonthStr = "0" + MonthStr
                '    End While
                '    RefNumber = RefNumber + "/" + MonthStr + "/" + vTanggalCDNote.Year.ToString
                'Else
                '    RefNumber = vCDNoteNumber
                'End If
                'If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                '    Prefix = "DNS"
                'Else
                '    Prefix = "CNS"
                'End If
                'RefNumber1 = RefNumber
                'While ObjJurnal.ApakahSudahAda(Prefix + RefNumber1) = True '-> memastikan tidak ada jurnal yang sama
                '    Counter += 1
                '    RefNumber1 = RefNumber + "(" + Counter.ToString + ")"
                'End While
                'RefNumber = RefNumber1

                If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                    RefNumber = "DNS" & Me.mNoJurnal 'RefNumber
                    'strDesc = "DEBIT NOTE TO SUPPLIER : " & vCDNoteNumber & " - " & SuppCode & " - " & SuppName & " - " & strRemarks
                ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    RefNumber = "CNS" & Me.mNoJurnal 'RefNumber
                    'strDesc = "CREDIT NOTE TO SUPPLIER : " & vCDNoteNumber & " - " & SuppCode & " - " & SuppName & " - " & strRemarks
                End If
                RefNumber = Me.mNoJurnal
                strDesc = Me.mDescJurnal
                'data header
                DTA_Jurnal.Insert_Journal(RefNumber, vTanggalCDNote, vUrutanCDNote, TipeJurnal, _
                                         vPrimaryCDNote, vPrimarySupp, strDesc, vUserCode, _
                                         Now, vUserCode, Now, Now, vTipeJurnal2)
                '                                                                                 '
                '---------------------------------------------------------------------------------'

                '---------------------------------------------------------------------------------'
                'data item
                PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, vTipeJurnal2, vPrimaryCDNote).Value
                NewPrimaryJournal = PrimaryJurnal
                '---------------------------------------------------------------------------------'
                ''simpan data vat in-debet
                'If vVAT <> 0 Then
                '    PrimaryAkun = Me.GetPrimaryAkunVATIn
                '    If PrimaryAkun <> -1 Then
                '        DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, PrimaryAkun, vVAT, 0, "", vVAT, 0, 1)
                '    End If
                'End If

                'simpan data ap, kredit
                PrimaryAkunAP = vPrimaryAkunAP
                PrimaryAkunAP_BaseCurrency = ObjCur.GetPrimaryAkunAP(Core_Module.PrimaryBaseCurrency)
                Dim PrimaryCurrReff As Double = -1
                PrimaryCurrReff = ObjAkun.GetPrimaryCurrency(vPrimaryAkunReff)
                Dim ExchangeRateReff As Decimal = 1
                If vNilaiCDNoteReff <> 0 Then
                    ExchangeRateReff = vNilaiCDNote / vNilaiCDNoteReff
                End If

                Debet = 0
                Credit = 0

                If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                    For Each vRowDgv In vDgv.Rows
                        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                            If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                                vRemarks = ""
                            Else
                                vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                            End If
                            With vRowDgv
                                If .Cells(vFrm.ColPrimaryCurrency.Index).Value = Core_Module.PrimaryBaseCurrency Then
                                    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                    .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                                Else
                                    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                    .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, _
                                    .Cells(vFrm.ColExchangeRate.Index).Value)
                                End If
                            End With
                        End If
                    Next
                ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    For Each vRowDgv In vDgv.Rows
                        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                            If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                                vRemarks = ""
                            Else
                                vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                            End If
                            With vRowDgv
                                If .Cells(vFrm.ColPrimaryCurrency.Index).Value = PrimaryAkunAP_BaseCurrency Then
                                    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                    .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                                Else
                                    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                    .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, _
                                    .Cells(vFrm.ColExchangeRate.Index).Value)
                                End If
                            End With
                        End If
                    Next
                End If

                ''buat purchase balance baru
                'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                '    'aop idr
                '    If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                '        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                '                                vCDNoteNumber, vTanggalCDNote, strRemarks, vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                '                                -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, _
                '                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '                                -vNilaiCDNote, 1, PrimaryAkunAP)


                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                '        vPrimarySupp, (vNilaiCDNote), 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                '        Core_Module.PrimaryBaseCurrency, (vNilaiCDNote), 0, 1, PrimaryAkunAP_BaseCurrency)

                '    ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                '        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                '                                vCDNoteNumber, vTanggalCDNote, strRemarks, vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                '                                vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, _
                '                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                '                                vNilaiCDNote, 1, PrimaryAkunAP)

                '        'buat kartu ap baru
                '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                '        vPrimarySupp, 0, (vNilaiCDNote), "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                '        Core_Module.PrimaryBaseCurrency, 0, (vNilaiCDNote), 1, PrimaryAkunAP_BaseCurrency)
                '    End If
                'Else
                '    If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                '        'balance purchase amount
                '        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                '        vCDNoteNumber, #1/1/1950#, strRemarks, vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                '        -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '        -vCDNoteOriginalAmount, vExchangeRate, _
                '        PrimaryAkunAP)

                '        'buat kartu ar purchase amount
                '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                '        vPrimarySupp, vNilaiCDNote, 0, _
                '        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                '        vPrimaryCurrency, vCDNoteOriginalAmount, 0, vExchangeRate, PrimaryAkunAP)

                '    ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                '        'balance purchase amount
                '        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                '        vCDNoteNumber, #1/1/1950#, strRemarks, vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                '        vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                '        vCDNoteOriginalAmount, vExchangeRate, _
                '        PrimaryAkunAP)

                '        'buat kartu ar purchase amount
                '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                '        vPrimarySupp, 0, vNilaiCDNote, _
                '        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                '        vPrimaryCurrency, 0, vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAP)
                '    End If
                'End If

                'buat purchase balance baru
                If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                    'aop idr
                    If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                                                vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                                                -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, _
                                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                                -vNilaiCDNote, 1, PrimaryAkunAP)


                        'buat kartu ap baru
                        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                        vPrimarySupp, (vNilaiCDNote), 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                        Core_Module.PrimaryBaseCurrency, (vNilaiCDNote), 0, 1, PrimaryAkunAP_BaseCurrency)

                    ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                                                vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                                                vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, _
                                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                                vNilaiCDNote, 1, PrimaryAkunAP)

                        'buat kartu ap baru
                        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                        vPrimarySupp, 0, (vNilaiCDNote), "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                        Core_Module.PrimaryBaseCurrency, 0, (vNilaiCDNote), 1, PrimaryAkunAP_BaseCurrency)
                    End If
                Else
                    If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                        vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                        -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        -vCDNoteOriginalAmount, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                        vPrimarySupp, vNilaiCDNote, 0, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                        vPrimaryCurrency, vCDNoteOriginalAmount, 0, vExchangeRate, PrimaryAkunAP)

                    ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                        'balance purchase amount
                        DTA_Balance.InsertQuery(vPrimarySupp, vPrimaryCDNote, TipeBalance, _
                        vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                        vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                        vCDNoteOriginalAmount, vExchangeRate, _
                        PrimaryAkunAP)

                        'buat kartu ar purchase amount
                        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                        vPrimarySupp, 0, vNilaiCDNote, _
                        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                        vPrimaryCurrency, 0, vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAP)
                    End If
                End If
            End If


                Return NewPrimaryJournal

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            DTAAkun.Dispose()
            ObjSupplier = Nothing
            ObjKartu = Nothing
            ObjSetting = Nothing
            ListAkunInv = Nothing
            ObjCD = Nothing
            ListPrimaryAkunCost = Nothing
        End Try
    End Function

    Public Function TransferJurnalCustomer(ByVal vPrimaryCDNote As Double, _
                                      ByVal vPrimaryCustomer As Double, _
                                      ByVal vNilaiCDNote As Decimal, _
                                      ByVal vNilaiCDNoteReff As Decimal, _
                                      ByVal vVAT As Decimal, _
                                      ByVal vUrutanCDNote As Double, _
                                      ByVal vTanggalCDNote As Date, _
                                      ByVal vCDNoteNumber As String, _
                                      ByVal vUserCode As String, _
                                      ByVal vPaymentTerm As Double, _
                                      ByVal vExchangeRate As Decimal, _
                                      ByVal vCDNoteOriginalAmount As Decimal, _
                                      ByVal vPrimaryCurrency As Double, _
                                      ByVal vPrimaryAkunReff As Double, _
                                      ByVal vPrimaryInvoice As Double, _
                                      ByVal vPrimaryAkunAP As Double, _
                                      ByVal TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2, _
                                      ByVal vTransaksiAsal As ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal, _
                                      ByVal vTipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi, ByVal strRemarks As String, ByVal vFrm As FrmTransaksiCreditNote) As Double
        '----------------------------------------------------------------------------------------------------------------------
        Dim DTA_Balance As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
        Dim DTA_Jurnal As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA_JurnalItem As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnal_itemTableAdapter
        Dim TipeJurnal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim RefNumber As String = CStr(vUrutanCDNote)
        Dim PrimaryJurnal As Double
        Dim Debet, Credit As Decimal
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjKartu As New ClsTransaksi_Kartu_APAR
        Dim MonthStr As String = ""
        Dim PrimaryAkunAR, PrimaryAkunAR_BaseCurrency As Double
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim PrimaryAkun As Double = -1
        Dim GabungDgnSales As Boolean = False
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DefPrecBase As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryBaseCurrency As Double = 1
        Dim ObjAkun As New ClsMasterAkun
        Dim PrimaryCurrencyAkunExp As Double = -1
        Dim NewPrimaryJournal As Double = -1
        Dim TipeNumbering As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
        Dim VATIDR As Decimal = 0
        Dim ObjJurnal As New ClsTransaksiJournal
        Dim Prefix As String = ""
        Dim Counter As Integer = 0
        Dim RefNumber1 As String = ""
        Dim vDgv As DataGridView = vFrm.dgvItem
        Dim vRowDgv As New DataGridViewRow
        Dim vRemarks As String = ""
        Try
            DefPrecBase = ObjCur.GetDecimal_DefaultCurrency
            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, TipeJurnal2, vPrimaryCDNote)

            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Credit_Note
            Else
                TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                TipeNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.Debit_Note
            End If

            'hapuskan data balance sales lama
            DTA_Balance.Delete(vPrimaryCDNote, TipeBalance)

            'hapus kartu lama
            ObjKartu.Hapus(vPrimaryCDNote, vTransaksiAsal)

            'siapkan nomor jurnal
            'If Me.ObjNumbering.Apakah_Autonumber(TipeNumbering) = True Then
            '    If RefNumber.Length < 6 Then
            '        While RefNumber.Length < 6
            '            RefNumber = "0" + RefNumber
            '        End While
            '    End If

            '    MonthStr = vTanggalCDNote.Month.ToString
            '    While MonthStr.Length < 2
            '        MonthStr = "0" + MonthStr
            '    End While

            '    RefNumber = RefNumber + "/" + MonthStr + "/" + vTanggalCDNote.Year.ToString
            'Else
            '    RefNumber = vCDNoteNumber
            'End If
            If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                Prefix = "DNC"
            Else
                Prefix = "CNC"
            End If
            RefNumber1 = RefNumber
            'While ObjJurnal.ApakahSudahAda(Prefix + RefNumber1) = True '-> memastikan tidak ada jurnal yang sama
            '    Counter += 1
            '    RefNumber1 = RefNumber + "(" + Counter.ToString + ")"
            'End While
            'RefNumber = RefNumber1

            'buat jurnal baru
            Dim strDesc As String = ""
            If vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                RefNumber = Me.mNoJurnal  '"DNC" & RefNumber
                'strDesc = "DEBIT NOTE FROM CUSTOMER : " & vCDNoteNumber & " - " & _
                '                                  ObjCust.GetCustomerCode(PrimaryCustomer) & _
                '                                  " - " + ObjCust.GetCustName(PrimaryCustomer) + " - " & strRemarks
            ElseIf vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                RefNumber = Me.mNoJurnal  '"CNC" & RefNumber
                'strDesc = "CREDIT NOTE FROM CUSTOMER : " & vCDNoteNumber & " - " & _
                '          ObjCust.GetCustomerCode(PrimaryCustomer) & " - " + ObjCust.GetCustName(PrimaryCustomer) + " - " & _
                '          strRemarks
            End If
            'RefNumber = Me.mNoJurnal
            strDesc = Me.mDescJurnal
            'data header
            DTA_Jurnal.Insert_Journal(RefNumber, vTanggalCDNote, _
            vUrutanCDNote, TipeJurnal, vPrimaryCDNote, PrimaryCustomer, strDesc, vUserCode, _
            Now, vUserCode, Now, Now, TipeJurnal2)


            'data item
            PrimaryJurnal = DTA_Jurnal.ScalarQuery_GetPrimaryJurnal(TipeJurnal, TipeJurnal2, vPrimaryCDNote).Value
            NewPrimaryJournal = PrimaryJurnal
            '======================================================================================================================================
            '======================================================================================================================================
            'Journal Account Receivable dulu
            '--------------------------------------------------------------------------------------------------------------------------------------
            'simpan data ar, debet
            PrimaryAkunAR = vPrimaryAkunAP
            PrimaryAkunAR_BaseCurrency = ObjCur.GetPrimaryAkunAR(1)
            Dim PrimaryCurrReff As Double = -1
            PrimaryCurrReff = ObjAkun.GetPrimaryCurrency(vPrimaryAkunReff)
            Dim Selisih As Decimal = 0
            Dim ExchangeRateInvoice As Decimal = 0
            Dim ExchangeRateReff As Decimal = 0

            If vNilaiCDNoteReff <> 0 Then
                ExchangeRateReff = vNilaiCDNote / vNilaiCDNoteReff
            End If

            'If vPrimaryInvoice > -1 Then
            '    ExchangeRateInvoice = Me.GetExchangeRateInvoice(vPrimaryInvoice)
            '    Selisih = (ExchangeRateInvoice - ExchangeRateReff) * vNilaiCDNoteReff
            'End If




            Debet = 0
            Credit = 0
            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                '=============================================================================================================================
                'SAVE Account Lawan
                'If PrimaryCurrReff = PrimaryBaseCurrency Then
                '    '-----------------------------------------------------------------------------------------------------------------------------
                '    Debet = vNilaiCDNote
                '    Credit = 0
                '    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, vPrimaryAkunReff, Debet, Credit, "", Debet, Credit, 1)
                '    '-----------------------------------------------------------------------------------------------------------------------------
                'Else
                '    '-----------------------------------------------------------------------------------------------------------------------------
                '    Debet = vNilaiCDNote
                '    Credit = 0
                '    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, vPrimaryAkunReff, Debet, Credit, "", vNilaiCDNoteReff, 0, ExchangeRateReff)
                '    '-----------------------------------------------------------------------------------------------------------------------------
                'End If
                ''=============================================================================================================================
                ''SAVE AR
                'If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '    '-----------------------------------------------------------------------------------------------------------------------------
                '    Debet = 0
                '    Credit = vNilaiCDNote
                '    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, Credit, 1)
                '    '-----------------------------------------------------------------------------------------------------------------------------
                'Else
                '    '-----------------------------------------------------------------------------------------------------------------------------
                '    Debet = 0
                '    Credit = vNilaiCDNote
                '    DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", 0, vCDNoteOriginalAmount, vExchangeRate)
                '    '-----------------------------------------------------------------------------------------------------------------------------
                'End If


                'Dim vRemarks As String = ""
                'For Each vRowDgv In vDgv.Rows
                '    If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                '        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                '            vRemarks = ""
                '        Else
                '            vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                '        End If
                '        With vRowDgv
                '            If .Cells(vFrm.ColPrimaryCurrency.Index).Value = PrimaryBaseCurrency Then
                '                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                '                .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                '            Else
                '                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                '                .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, _
                '                .Cells(vFrm.ColExchangeRate.Index).Value)
                '            End If
                '        End With
                '    End If
                'Next
                For Each vRowDgv In vDgv.Rows
                    If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                            vRemarks = ""
                        Else
                            vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                        End If
                        With vRowDgv
                            If .Cells(vFrm.ColPrimaryCurrency.Index).Value = PrimaryBaseCurrency Then
                                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                            Else
                                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, _
                                .Cells(vFrm.ColExchangeRate.Index).Value)
                            End If
                        End With
                    End If
                Next
                '=============================================================================================================================
            ElseIf vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                '=============================================================================================================================
                '    'SAVE AR
                '    If PrimaryAkunAR = PrimaryAkunAR_BaseCurrency Then
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '        Debet = vNilaiCDNote
                '        Credit = 0
                '        DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", Debet, Credit, 1)
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        Debet = vNilaiCDNote
                '        Credit = 0
                '        DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, PrimaryAkunAR, Debet, Credit, "", vCDNoteOriginalAmount, 0, vExchangeRate)
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '    End If

                '    '=============================================================================================================================
                '    'SAVE Account Lawan
                '    If PrimaryCurrReff = PrimaryBaseCurrency Then
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '        Debet = 0
                '        Credit = vNilaiCDNote
                '        DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, vPrimaryAkunReff, Debet, Credit, "", Debet, Credit, 1)
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '    Else
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '        Debet = 0
                '        Credit = vNilaiCDNote
                '        DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, vPrimaryAkunReff, Debet, Credit, "", 0, vNilaiCDNoteReff, ExchangeRateReff)
                '        '-----------------------------------------------------------------------------------------------------------------------------
                '    End If
                'End If
                'Dim vRemarks As String = 0
                'For Each vRowDgv In vDgv.Rows
                '    If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                '        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                '            vRemarks = ""
                '        Else
                '            vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                '        End If
                '        With vRowDgv
                '            If .Cells(vFrm.ColPrimaryCurrency.Index).Value = PrimaryBaseCurrency Then
                '                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                '                  .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                '            Else
                '                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, .Cells(vFrm.ColKredit.Index).Value, _
                '                vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, .Cells(vFrm.ColExchangeRate.Index).Value)
                '            End If
                '        End With
                '    End If
                'Next
                For Each vRowDgv In vDgv.Rows
                    If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColAccount.Index).Value) = False Then
                        If String.IsNullOrEmpty(vRowDgv.Cells(vFrm.ColRemarks.Index).Value) = True Then
                            vRemarks = ""
                        Else
                            vRemarks = vRowDgv.Cells(vFrm.ColRemarks.Index).Value
                        End If
                        With vRowDgv
                            If .Cells(vFrm.ColPrimaryCurrency.Index).Value = PrimaryBaseCurrency Then
                                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, _
                                  .Cells(vFrm.ColKredit.Index).Value, vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, 1)
                            Else
                                DTA_JurnalItem.Insert_Journal_Item(PrimaryJurnal, .Cells(vFrm.ColPrimaryAkun.Index).Value, .Cells(vFrm.ColDebet.Index).Value, .Cells(vFrm.ColKredit.Index).Value, _
                                vRemarks, .Cells(vFrm.ColDebet_OriginalCurr.Index).Value, .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value, .Cells(vFrm.ColExchangeRate.Index).Value)
                            End If
                        End With
                    End If
                Next
            End If



            ''buat purchase balance baru
            'If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
            '    'aop idr
            '    If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
            '        DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
            '                                                    vCDNoteNumber, vTanggalCDNote, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
            '                                                    -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, _
            '                                                    Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
            '                                                    -vNilaiCDNote, 1, PrimaryAkunAR)


            '        'buat kartu ap baru
            '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
            '        vPrimaryCustomer, 0, (vNilaiCDNote), "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
            '        Core_Module.PrimaryBaseCurrency, 0, (vNilaiCDNote), 1, PrimaryAkunAR_BaseCurrency)

            '    ElseIf vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
            '        DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
            '                                vCDNoteNumber, vTanggalCDNote, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
            '                                vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, _
            '                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
            '                                vNilaiCDNote, 1, PrimaryAkunAR)

            '        'buat kartu ap baru
            '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
            '        vPrimaryCustomer, (vNilaiCDNote), 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
            '        Core_Module.PrimaryBaseCurrency, (vNilaiCDNote), 0, 1, PrimaryAkunAR_BaseCurrency)
            '    End If
            'Else
            '    If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then

            '        'balance purchase amount
            '        DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
            '         vCDNoteNumber, vTanggalCDNote, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
            '        -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
            '        -vCDNoteOriginalAmount, vExchangeRate, _
            '        PrimaryAkunAR)

            '        'buat kartu ar purchase amount
            '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
            '        vPrimaryCustomer, 0, vNilaiCDNote, _
            '        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
            '        vPrimaryCurrency, 0, vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAR)


            '    ElseIf vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
            '        'balance purchase amount
            '        DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
            '        vCDNoteNumber, vTanggalCDNote, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
            '        vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
            '        vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAR)

            '        'buat kartu ar purchase amount
            '        ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
            '        vPrimaryCustomer, vNilaiCDNote, 0, _
            '        "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
            '        vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
            '        vPrimaryCurrency, vCDNoteOriginalAmount, 0, vExchangeRate, PrimaryAkunAR)
            '    End If

            '    ''tax amount + duty stamp
            '    'If vVAT + AP_PadaCost_Base <> 0 Then
            '    '    'balance tax amount + duty stamp
            '    '    DTA_Balance.InsertQuery(vPrimaryCustomerCust, vPrimaryCDNote, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
            '    '    0, vVAT + AP_PadaCost_Base, 0, vVAT, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, vVAT, 1, _
            '    '    PrimaryAkunAR_BaseCurrency)

            '    '    'buat kartu ar tax amount + duty stamp
            '    '    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
            '    '    vPrimaryCustomerCust, 0, vVAT, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
            '    '    vTanggalAOP, vPrimaryCDNote, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
            '    '    Core_Module.PrimaryBaseCurrency, 0, vVAT, PrimaryAkunAR_BaseCurrency, 1)
            '    'End If
            'End If

            'buat purchase balance baru
            If vPrimaryCurrency = Core_Module.PrimaryBaseCurrency Then
                'aop idr
                If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
                                                                vCDNoteNumber, #1/1/1950#, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                                                                -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, _
                                                                Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                                                -vNilaiCDNote, 1, PrimaryAkunAR)


                    'buat kartu ap baru
                    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                    vPrimaryCustomer, 0, (vNilaiCDNote), "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                    Core_Module.PrimaryBaseCurrency, 0, (vNilaiCDNote), 1, PrimaryAkunAR_BaseCurrency)

                ElseIf vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                    DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
                                            vCDNoteNumber, #1/1/1950#, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                                            vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, _
                                            Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, _
                                            vNilaiCDNote, 1, PrimaryAkunAR)

                    'buat kartu ap baru
                    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                    vPrimaryCustomer, (vNilaiCDNote), 0, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                    Core_Module.PrimaryBaseCurrency, (vNilaiCDNote), 0, 1, PrimaryAkunAR_BaseCurrency)
                End If
            Else
                If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then

                    'balance purchase amount
                    'DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
                    ' vCDNoteNumber, #1/1/1950#, "", vTanggalCDNote + TimeSpan.FromDays(vPaymentTerm), _
                    '-vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                    '-vCDNoteOriginalAmount, vExchangeRate, _
                    'PrimaryAkunAR)
                    DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
                   vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                  -vNilaiCDNote, 0, 0, -vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                  -vCDNoteOriginalAmount, vExchangeRate, _
                  PrimaryAkunAR)

                    'buat kartu ar purchase amount
                    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, vCDNoteNumber, _
                    vPrimaryCustomer, 0, vNilaiCDNote, _
                    "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note, _
                    vPrimaryCurrency, 0, vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAR)


                ElseIf vTipeTransaksi = EnumTipeTransaksi.DEBIT_NOTE Then
                    'balance purchase amount
                    DTA_Balance.InsertQuery(vPrimaryCustomer, vPrimaryCDNote, TipeBalance, _
                    vCDNoteNumber, #1/1/1950#, "", Me.mDueDate, _
                    vNilaiCDNote, 0, 0, vNilaiCDNote, 0, vUserCode, Now, vUserCode, Now, Now, vPrimaryCurrency, _
                    vCDNoteOriginalAmount, vExchangeRate, PrimaryAkunAR)

                    'buat kartu ar purchase amount
                    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, vCDNoteNumber, _
                    vPrimaryCustomer, vNilaiCDNote, 0, _
                    "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AR, _
                    vTanggalCDNote, vPrimaryCDNote, vTanggalCDNote, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note, _
                    vPrimaryCurrency, vCDNoteOriginalAmount, 0, vExchangeRate, PrimaryAkunAR)
                End If


                ''tax amount + duty stamp
                'If vVAT + AP_PadaCost_Base <> 0 Then
                '    'balance tax amount + duty stamp
                '    DTA_Balance.InsertQuery(vPrimaryCustomerCust, vPrimaryCDNote, TipeBalance, "", #1/1/1950#, "", vTanggalAOP + TimeSpan.FromDays(vPaymentTerm), _
                '    0, vVAT + AP_PadaCost_Base, 0, vVAT, 0, vUserCode, Now, vUserCode, Now, Now, Core_Module.PrimaryBaseCurrency, vVAT, 1, _
                '    PrimaryAkunAR_BaseCurrency)

                '    'buat kartu ar tax amount + duty stamp
                '    ObjKartu.Simpan(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, vAOPNumber, _
                '    vPrimaryCustomerCust, 0, vVAT, "-", ClsTransaksi_Kartu_APAR.EnumJenisKartu.AP, _
                '    vTanggalAOP, vPrimaryCDNote, vTanggalAOP, vPaymentTerm, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.AOP_Otomatis, _
                '    Core_Module.PrimaryBaseCurrency, 0, vVAT, PrimaryAkunAR_BaseCurrency, 1)
                'End If
            End If

            Return NewPrimaryJournal

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            DTA_JurnalItem.Dispose()
            ObjCust = Nothing
            ObjKartu = Nothing
            ObjCur = Nothing
            ObjSetting = Nothing
            ObjCD = Nothing
        End Try

    End Function

    Private Function ApakahTransferOtomatis(ByVal vTanggalInvoice As Date) As Boolean
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_accounting_closingTableAdapter

        Try
            'Transfer otomatis jika:
            '1. Tanggal invoice >= convertion date
            If DTA.ScalarQuery_GetConvertionDate.HasValue = True Then
                If DTA.ScalarQuery_GetConvertionDate.Value.Date <= vTanggalInvoice.Date Then
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


    Private Function GetPrimaryAkunVATIn() As Double
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Dummy As String = ""

        Try
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
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


    Private Function GetPrimaryDutyStamp() As Double
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

    'Private Function GetPrimaryAkunAP(ByVal vPrimaryCDNote As Double) As Double
    '    Dim ObjCur As New Core_Module.clsMasterCurrency
    '    Dim ObjInv As New ClsTransaksiSalesInvoice

    '    Try
    '        Return ObjCur.GetPrimaryAkunAP(ObjInv.Get_Primary_Currency(vPrimaryCDNote))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjCur = Nothing
    '        ObjInv = Nothing
    '    End Try
    'End Function

    'Private Function GetPrimaryAkunAR(ByVal vPrimaryCDNote As Double) As Double
    '    'Dim ObjCur As New Core_Module.clsMasterCurrency
    '    'Dim ObjAOP As New ClsTransaksiAOP

    '    Try
    '        Return ObjCur.GetPrimaryAkunAR(ObjAOP.GetPrimaryCurrency(vPrimaryCDNote))
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        'ObjCur = Nothing
    '        'ObjAOP = Nothing
    '    End Try
    'End Function

    Private Function GetExchangeRateInvoice(ByVal vPrimaryInvoice As Double) As Decimal
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_salesinvoiceTableAdapter
        Dim decExchangeRate As Decimal = 0
        Try
            If DTA.ScalarQuery_GetExchangerate(vPrimaryInvoice).HasValue AndAlso DTA.ScalarQuery_GetExchangerate(vPrimaryInvoice).Value > 0 Then
                decExchangeRate = CDec(DTA.ScalarQuery_GetExchangerate(vPrimaryInvoice).Value)
            End If

            Return decExchangeRate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetExchangeRateAOP(ByVal vPrimaryAOP As Double) As Decimal
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_aopTableAdapter
        Dim decExchangeRate As Decimal = 0
        Try
            If DTA.ScalarQuery_GetExchangeRate(vPrimaryAOP).HasValue AndAlso DTA.ScalarQuery_GetExchangeRate(vPrimaryAOP).Value > 0 Then
                decExchangeRate = CDec(DTA.ScalarQuery_GetExchangeRate(vPrimaryAOP).Value)
            End If

            Return decExchangeRate
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Sub Hapus_Jurnal(ByVal vPrimaryCDNote As Double, ByVal vTipeTransaksi As EnumTipeTransaksi, ByVal VTipeKartu As EnumTipeKartu)
        Dim DTA_Balance As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTA_Balance1 As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim DTA_Jurnal As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance
        Dim TipeBalance2 As ClsTransaksiAOP.EnumTipePurchaseBalance
        Dim TipeJurnal As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Credit_Note
        Dim Obj_KartuAR As New ClsTransaksi_Kartu_APAR

        Try
            DTA_Balance.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Balance1.Connection.ConnectionString = DTA_Balance.Connection.ConnectionString
            DTA_Jurnal.Connection.ConnectionString = DTA_Balance.Connection.ConnectionString

            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Credit_Note
                TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                TipeBalance2 = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
            Else
                TipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.Debit_Note
                TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                TipeBalance2 = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
            End If

            'hapuskan data jurnal lama
            DTA_Jurnal.Delete(TipeJurnal, TipeJurnal2, vPrimaryCDNote)

            'hapuskan data balance sales lama
            If VTipeKartu = EnumTipeKartu.Customer Then
                DTA_Balance.Delete(vPrimaryCDNote, TipeBalance)
            ElseIf VTipeKartu = EnumTipeKartu.Supplier Then
                DTA_Balance1.Delete(vPrimaryCDNote, TipeBalance2)
            End If

            'hapuskan kartu AR lama
            If vTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                Obj_KartuAR.Hapus(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Credit_Note)
            Else
                Obj_KartuAR.Hapus(vPrimaryCDNote, ClsTransaksi_Kartu_APAR.EnumJenisTransaksiAsal.Debit_Note)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA_Balance1.Dispose()
            DTA_Balance.Dispose()
            DTA_Jurnal.Dispose()
            Obj_KartuAR = Nothing
        End Try
    End Sub
#End Region

    'Public Function ApakahBolehHapus(ByVal vPrimary As Double) As Boolean

    '    Try

    '        Return True
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function
    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
          ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, ByVal vPrimaryJournal As Double, ByVal vTipe As Integer, _
          Optional ByVal TampilkanPesan As Boolean = True) As Boolean

        Dim DTAJournal As New DseTransaksiJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjClosing As New ClsTransaksiClosing
        Dim TipeJurnal1 As ClsTransaksiJournal.EnumTipeJournal = ClsTransaksiJournal.EnumTipeJournal.GeneralJournal
        Dim TipeJurnal2 As ClsTransaksiJournal.EnumTipeJurnal2 = ClsTransaksiJournal.EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation
        Dim PrimaryLastJurnal As Double = -1
        'Dim LastRevaluationDate, DummyDate As Date
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim NoClosing As String = ""
        Dim ObjJournal As New ClsTransaksiJournal
        Dim DTAXSales As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTAXPurchase As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim Acuan As String = ""
        Dim TipeBalance As ClsTransaksiSalesBalance.EnumTipeSalesBalance
        Dim TglCNDN As Date = Today
        Dim ObjUser As New Core_Module.ClsUser
        Dim TipeTrans As EnumTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAJournal.Connection.ConnectionString = Core_Module.GetConstring
            DTA_Closing.Connection.ConnectionString = DTAJournal.Connection.ConnectionString
            DTAXSales.Connection.ConnectionString = DTAJournal.Connection.ConnectionString
            DTAXPurchase.Connection.ConnectionString = DTAJournal.Connection.ConnectionString

            '1. jika tanggal sebelum tanggal closing maka tidak bisa dihapus
            If vPrimaryJournal > -1 Then
                If ObjJournal.GetTanggalJournal(vPrimaryJournal).Date <= ObjClosing.GetClosingDate.Date Then
                    If TampilkanPesan = True Then
                        MsgBox("You cannot delete this data because transaction journal date <= closing date " _
                        , MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                TipeTrans = Me.GetTipeTransksi(vPrimaryTransaksi)
                If TipeTrans = EnumTipeTransaksi.CREDIT_NOTE Then
                    If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiCreditNote_HapusUserLain") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                ElseIf TipeTrans = EnumTipeTransaksi.DEBIT_NOTE Then
                    If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiDebetNote_HapusUserLain") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If

                'If TampilkanPesan = True Then
                '    MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                '    MsgBoxStyle.Information, "Information")
                'End If
                'Return False
            End If

            TipeBalance = vTipe
            If DTAXSales.ScalarQuery_getReceived(vPrimaryTransaksi, vTipe).HasValue AndAlso DTAXSales.ScalarQuery_getReceived(vPrimaryTransaksi, vTipe).Value <> 0 Then
                Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, vPrimaryTransaksi)
                MsgBox("You cannot delete this data because transaction journal already received on sales receivement journal: " & Acuan, MsgBoxStyle.Information, "Information")
                Return False
            End If

            If DTAXPurchase.ScalarQuery_GetPaid(vPrimaryTransaksi, vTipe).HasValue AndAlso DTAXPurchase.ScalarQuery_GetPaid(vPrimaryTransaksi, vTipe).Value <> 0 Then
                Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, vPrimaryTransaksi)
                MsgBox("You cannot delete this data because transaction journal already payed on purchase payment journal: " & Acuan, MsgBoxStyle.Information, "Information")
                Return False
            End If

            '3. cek ke transaksi custom
            'Dim ObjCust As New CustomSMESS.ClsTransaksiJournal
            'If ObjCust.ApakahBolehHapusCustom(vPrimaryTransaksi) = False Then
            '    Return False
            'End If
            'ObjCust = Nothing

            '4. jika tipenya adalah end of month journal, maka hanya bisa diedit untuk transaksi yang terakhir saja
            'If DTAJournal.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).HasValue = True Then

            '    'TipeJurnal2 = DTAJournal.ScalarQuery_GetTipeJurnal2(vPrimaryTransaksi).Value
            '    'If TipeJurnal2 = EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation Then
            '    '    TipeJurnal1 = EnumTipeJournal.GeneralJournal

            '    '    If DTAJournal.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).HasValue = True Then
            '    '        PrimaryLastJurnal = DTAJournal.ScalarQuery_GetPrimaryMaxByTipeJurnal2(TipeJurnal1, TipeJurnal2).Value
            '    '        If PrimaryLastJurnal <> vPrimaryTransaksi Then
            '    '            MsgBox("You cannot delete this data, because it's not the last period of end of month journal.", _
            '    '            MsgBoxStyle.Information, "Information")
            '    '            Return False
            '    '        End If
            '    '    End If
            '    'End If

            '    ''5. jika membuat jurnal umum, selain jurnal revaluasi, tidak boleh membuat jurnal yg memiliki mata uang non IDR
            '    ''pada periode yang sudah direvaluasi
            '    'TipeJurnal1 = Me.GetTipeJournal(vPrimaryTransaksi)
            '    'LastRevaluationDate = ObjCD.GetLasDateOfTheMonth(Me.GetTglRevaluasiSelanjutnya - TimeSpan.FromDays(33))
            '    'DummyDate = ObjCD.GetFirstDateofTheMonth(Me.GetTanggalJournal(vPrimaryTransaksi))
            '    'If (TipeJurnal1 = EnumTipeJournal.GeneralJournal And TipeJurnal2 <> EnumTipeJurnal2.GeneralJurnal_MonthlyRevaluation) Or _
            '    'TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Or TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Then
            '    '    If LastRevaluationDate >= DummyDate Then
            '    '        If Me.ApakaJurnalNonIDR(vPrimaryTransaksi) = True Then
            '    '            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
            '    '                MsgBoxStyle.Information, "Information")
            '    '            Return False
            '    '        End If
            '    '    End If
            '    'End If
            '    'If TipeJurnal1 = EnumTipeJournal.SalesReceivementJournal Or TipeJurnal1 = EnumTipeJournal.PurchasePaymentJournal Then
            '    '    If LastRevaluationDate >= DummyDate Then
            '    '        If Me.ApakahAdaPaymentNonIDR(vPrimaryTransaksi) = True Then
            '    '            MsgBox("The period of " + Format(DummyDate, "MMMMM yyyy") + " has been revaluate, you cannot edit this data", _
            '    '                MsgBoxStyle.Information, "Information")
            '    '            Return False
            '    '        End If
            '    '    End If
            '    'End If
            'End If

            ''jika jurnal tersebut dirujuk oleh purchase closing
            'If vPrimaryJournal > -1 Then
            '    If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryJournal)) = False Then
            '        NoClosing = DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryJournal)
            '    Else
            '        NoClosing = ""
            '    End If
            '    If NoClosing <> "" Then
            '        MsgBox("This transaction is referred by inventory closing: " + NoClosing + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            If Me.GetPrimaryCurrency(vPrimaryJournal) <> Core_Module.PrimaryBaseCurrency Then
                TglCNDN = Me.GetTanggal(vPrimaryTransaksi)
                If ObjJournal.ApakahSudahRevaluasi(TglCNDN) = True Then
                    MsgBox("There's end of month journal for period " + Format(TglCNDN, "MMM yyyy") + _
                    ", you can't delete credit note/debit note with non-IDR currency.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAJournal.Dispose()
            DTA_Closing.Dispose()
            DTAXSales.Dispose()
            DTAXPurchase.Dispose()
            ObjClosing = Nothing
            ObjCD = Nothing
            ObjJournal = Nothing
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double, ByVal VTipeTransaksi As EnumTipeTransaksi, ByVal vTipeKartu As EnumTipeKartu, _
    ByVal vPrimaryJournal As Double, ByVal vUserCode As String) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim ObjAntri As New Core_Module.ClsAntrian
        Dim SW As New Stopwatch
        Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter

        Try
            Core_Module.SetConstring_DTAAntrian(DTAInOUT)

            SW.Start()

            'ambil no antrian
            ObjAntri.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, DTAInOUT, Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice, _
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.CreditDebitNote_Hapus, vUserCode)


            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If SW.Elapsed.TotalSeconds >= ObjAntri.GetWaitTime Then
                        MsgBox("Database is busy please try again in a few moments.", MsgBoxStyle.Information, "Information")
                        DTAInOUT.Delete(Me.vPrimary_Antrian)
                        Exit Function
                    Else
                        If DTAInOUT.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                            'clock antri matikan
                            Me.Clock_Antri.Stop()
                            'sudah saatnya melakukan proses simpan baru
                            retval = Me.Hapus1(vPrimary, VTipeTransaksi, vTipeKartu, vPrimaryJournal, vUserCode)
                            'hapus data antri untuk transaksi ini
                            If DTAInOUT.Delete(Me.vPrimary_Antrian) = 0 Then
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

    Public Function Hapus1(ByVal vPrimary As Double, ByVal VTipeTransaksi As EnumTipeTransaksi, ByVal vTipeKartu As EnumTipeKartu, ByVal vPrimaryJournal As Double, ByVal vUserCode As String) As Boolean

        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dim TipeBalance As Integer = 0
            If vTipeKartu = EnumTipeKartu.Supplier Then
                If VTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Credit_Note
                Else
                    TipeBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Debet_Note
                End If
            Else
                If VTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                    TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Credit_Note
                Else
                    TipeBalance = ClsTransaksiSalesBalance.EnumTipeSalesBalance.Debet_Note
                End If
            End If

            If Me.ApakahBolehHapus(vPrimary, True, vUserCode, vPrimaryJournal, TipeBalance, False) = True Then
                Hapus_Jurnal(vPrimary, VTipeTransaksi, vTipeKartu)
                DTA.Delete(CDec(vPrimary))
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

    Public Function ApakahSudahAda(ByVal vNoTrans As String) As Boolean
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vNoTrans).HasValue Then
                If DTA.ScalarQuery_GetJml(vNoTrans).Value > 0 Then
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

    Public Function IsiDGV(ByVal Criteria As String, _
   ByVal DGV As DataGridView, ByVal TBL As DseTransaksiCreditNote_TransferJournal.tbltrans_credit_note_viewDataTable, _
   ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
     ByVal vTipeTransaksi As ClsTransaksiCreditNote.EnumTipeTransaksi, ByVal vTanggalDari As Date, _
     ByVal vTanggalSampai As Date, ByVal FilterCreatedBy As String) As Boolean
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_credit_note_viewTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Dim Reset As Integer = 0
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
            '    QueryResult = DTA.Fill(TBL)
            'End If
            '''' 
            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_LikeNoJurnal(TBL, vTipeTransaksi, Reset, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDate Then
                QueryResult = DTA.FillBy_tglSampaiDari(TBL, vTanggalSampai, vTanggalDari, vTipeTransaksi, Reset)
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_LikeReason(TBL, Criteria + "%", vTipeTransaksi, Reset)
            ElseIf SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTrans(TBL, Criteria + "%", vTipeTransaksi, Reset)
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescJurnal Then
                QueryResult = DTA.FillBy_DescJurnalLIKE(TBL, Criteria + "%", Reset, vTipeTransaksi)
            Else
                QueryResult = DTA.FillBy_reference(TBL, vTipeTransaksi, Criteria + "%", Reset)
            End If

            'update DGV
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If FilterCreatedBy <> "" Then
                TBL.DefaultView.RowFilter = FilterCreatedBy
                QueryResult = TBL.DefaultView.Count
            Else
                TBL.DefaultView.RowFilter = ""
            End If
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(1).HeaderText = "Document No"
                DGV.Columns(2).HeaderText = "Date"
                DGV.Columns(3).HeaderText = "Journal No"
                DGV.Columns(4).HeaderText = "Journal Description"
                DGV.Columns(5).HeaderText = "Reference Party"
                DGV.Columns(6).HeaderText = "Reason"
                DGV.Columns(1).MinimumWidth = 130
                DGV.Columns(2).MinimumWidth = 130
                DGV.Columns(3).MinimumWidth = 130
                DGV.Columns(4).MinimumWidth = 130
                DGV.Columns(5).MinimumWidth = 130
                DGV.Columns(6).MinimumWidth = 130
                DGV.Columns(7).MinimumWidth = 130
                DGV.Columns(7).Visible = False
                DGV.Columns(8).Visible = False
                DGV.Columns(9).Visible = False
                DGV.Columns(10).Visible = False
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.TotalSeconds > 1 Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + objCD.DigitGrouping(TS.TotalSeconds, 2) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, 0, MidpointRounding.AwayFromZero)) + _
                    " miliseconds"
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

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
    ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, ByVal vPrimaryJournal As Double, ByVal vTipe As Integer, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim DTAXSales As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_salesTableAdapter
        Dim DTAXPurchase As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_balance_purchaseTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA_Closing As New DseTransaksiMonthlyClosingTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim NoClosing As String = ""
        Dim TipeBalance As ClsTransaksiAOP.EnumTipePurchaseBalance = ClsTransaksiAOP.EnumTipePurchaseBalance.Automatic
        Dim Acuan As String = ""
        Dim ObjJournal As New ClsTransaksiJournal
        Dim TglCNDN As Date = Today
        Dim ObjUser As New Core_Module.ClsUser
        Dim TipeTrans As EnumTipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTAXSales.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAXPurchase.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTA_Closing.Connection.ConnectionString = DTA.Connection.ConnectionString

            'jika strict Editting maka hanya sales staff yang bolEh edit
            If StrictEditting = True Then
                CreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    TipeTrans = Me.GetTipeTransksi(vPrimaryTransaksi)
                    If TipeTrans = EnumTipeTransaksi.CREDIT_NOTE Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiCreditNote_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    ElseIf TipeTrans = EnumTipeTransaksi.DEBIT_NOTE Then
                        If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiDebetNote_EditUserLain") = False Then
                            If TampilkanPesan = True Then
                                MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                                MsgBoxStyle.Information, "Information")
                            End If
                            Return False
                        End If
                    End If
                End If

                'If CreatedBy <> vSalesStaff Then
                '    If TampilkanPesan = True Then
                '        MsgBox("Only user with ID : " + CreatedBy + " can edit this data.", MsgBoxStyle.Information, "Information")
                '    End If
                '    Return False
                'End If
            End If

            TipeBalance = vTipe
            If DTAXSales.ScalarQuery_getReceived(vPrimaryTransaksi, vTipe).HasValue AndAlso CDec(DTAXSales.ScalarQuery_getReceived(vPrimaryTransaksi, vTipe).Value) <> 0 Then
                'cari sales receivement journal yang merujuk pada transaksi ini
                Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, vPrimaryTransaksi)
                MsgBox("You cannot edit this data because transaction journal already received on sales receivement journal: " + Acuan, MsgBoxStyle.Information, "Information")
                Return False
            End If

            If DTAXPurchase.ScalarQuery_GetPaid(vPrimaryTransaksi, vTipe).HasValue AndAlso CDec(DTAXPurchase.ScalarQuery_GetPaid(vPrimaryTransaksi, vTipe).Value) <> 0 Then
                Acuan = ObjJournal.GetNoTransaksi_ByDocumentAcuan(TipeBalance, vPrimaryTransaksi)
                MsgBox("You cannot edit this data because transaction journal already payed on purchase payment journal: " + Acuan, MsgBoxStyle.Information, "Information")
                Return False
            End If

            ''jika jurnal tersebut dirujuk oleh purchase closing
            'If vPrimaryJournal > -1 Then
            '    If String.IsNullOrEmpty(DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryJournal)) = False Then
            '        NoClosing = DTA_Closing.ScalarQuery_GetNoTransaksi(vPrimaryJournal)
            '    Else
            '        NoClosing = ""
            '    End If
            '    If NoClosing <> "" Then
            '        MsgBox("This transaction referred by inventory closing: " + NoClosing + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
            '        Return False
            '    End If
            'End If

            If Me.GetPrimaryCurrency(vPrimaryJournal) <> Core_Module.PrimaryBaseCurrency Then
                TglCNDN = Me.GetTanggal(vPrimaryTransaksi)
                If ObjJournal.ApakahSudahRevaluasi(TglCNDN) = True Then
                    MsgBox("There's end of month journal for period " + Format(TglCNDN, "MMM yyyy") + _
                    ", you can't edit credit note/debit note with non-IDR currency.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'sampai sini balikan nilai true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTAXSales.Dispose()
            DTAXPurchase.Dispose()
            DTA_Closing.Dispose()
            ObjJournal = Nothing
        End Try
    End Function

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim RetVal As String = ""
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetCreatedBy(vPrimary)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetTipeTransksi(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As Nullable(Of Integer)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetTipeTransaksi(vPrimary)
            If Retval.HasValue = True Then
                Return Retval.Value
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

    Public Function GetTanggal(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_credit_note_viewTableAdapter
        Dim retVal As Nullable(Of Date)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retVal = DTA.ScalarQuery_GetTanggalTransaksi(vPrimary)
            If retVal.HasValue = True Then
                Return retVal.Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return #1/1/1950#
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryTrans As Double) As Double
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_credit_note_viewTableAdapter
        Dim retVal As Nullable(Of Double)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            retVal = DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryTrans)
            If retVal.HasValue = True Then
                Return retVal.Value
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

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vFrm As FrmTransaksiCreditNote) As Boolean
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_note2TableAdapter
        Dim TBL As New DseTransaksiCreditNote.tbltrans_credit_note2DataTable
        Dim row As DseTransaksiCreditNote.tbltrans_credit_note2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, CDec(vPrimary))
            If TBL.Rows.Count > 0 Then
                row = TryCast(TBL.Rows(0), DseTransaksiCreditNote.tbltrans_credit_note2Row)
                With row
                    Me.mPrimary = row.primary
                    Me.mPrimaryCustomer = row.primary_customer
                    Me.mPrimaryAkunAr = row.primary_akun_ar
                    Me.mPrimaryExpenseAcc = row.expense_account
                    Me.mPrimaryCurrency = row.primary_currency
                    Me.mPrimaryInvoice = row.primary_invoice
                    Me.mPrimaryJournal = row.primary_jurnal
                    Me.mTgl = row.tgl
                    Me.mUrutan = row.urutan
                    Me.mNoTrans = row.no_transksi
                    Me.mIdentifierAddCust = row.identifier_add_cust
                    Me.mCustomerDebitNote = row.customer_debet_note
                    Me.mAmount = row.amount
                    Me.mExchangeRate = row.Exchange_Rate
                    Me.mDueDate = row.due_date
                    Me.mReason = row.reason
                    Me.mTipeTransaksi = row.tipe_transaksi
                    Me.mTipeKartu = row.tipe_kartu
                    Me.mContactPerson = row.contact_person
                    Me.IsBeginingBalance = row.is_begining_balance
                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    Me.mAmountReff = .Amount_Reff
                    Me.mNoJurnal = .no_jurnal
                    Me.mDescJurnal = .description_jurnal
                End With
            End If

            Dim ObjCurr As New Core_Module.clsMasterCurrency
            Dim Curr As String = ObjCurr.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)
            Dim DefPrec As Decimal = ObjCurr.GetDefaultPrecission(Core_Module.PrimaryBaseCurrency)
            IsiDgvItem(mPrimaryJournal, mPrimary, mPrimaryCurrency, Curr, vFrm, DefPrec)
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Private Function IsiDgvItem(ByVal vPrimaryJurnal As Double, ByVal vPrimaryDocAcuan As Double, ByVal vPrimaryCurr As Double, ByVal vCurrCode As String, _
   ByVal vFrm As FrmTransaksiCreditNote, ByVal vDefPrec As Decimal) As Boolean
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiCreditNote.DataTable2DataTable
        Dim Row As DseTransaksiCreditNote.DataTable2Row
        Dim Dgv As DataGridView = vFrm.dgvItem
        Dim RowDgv As New DataGridViewRow
        Try
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Dgv.Rows.Clear()
            Dgv.AllowUserToAddRows = False
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Fill(TBL, vPrimaryJurnal, vPrimaryDocAcuan) > 0 Then
                For Each Row In TBL.Rows
                    Dgv.Rows.Add()
                    RowDgv = Dgv.Rows(Dgv.Rows.Count - 1)
                    With RowDgv
                        .Cells(vFrm.ColPrimaryAkun.Index).Value = Row.Primary_Akun
                        .Cells(vFrm.ColPrimaryCurrency.Index).Value = Row.primary_currency
                        .Cells(vFrm.ColOriginalCurrency.Index).Value = Row.Currency_Code
                        .Cells(vFrm.ColAccount.Index).Value = Row.Account_Code
                        .Cells(vFrm.ColDescAccount.Index).Value = Row.Description
                        .Cells(vFrm.ColCredit_OriginalCurrency.Index).Value = Me.objCD.DigitGrouping(Row.Credit_Original, Row.Default_Precision, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        .Cells(vFrm.ColDebet_OriginalCurr.Index).Value = Me.objCD.DigitGrouping(Row.Debet_Original, Row.Default_Precision, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        .Cells(vFrm.ColDebet.Index).Value = Me.objCD.DigitGrouping(Row.Debet, vDefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        .Cells(vFrm.ColKredit.Index).Value = Me.objCD.DigitGrouping(Row.Kredit, vDefPrec, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        .Cells(vFrm.ColPrimaryDept.Index).Value = Row.primary_dept
                        .Cells(vFrm.ColDefPrec.Index).Value = Row.Default_Precision
                        If .Cells(vFrm.ColPrimaryCurrency.Index).Value = Core_Module.PrimaryBaseCurrency Then
                            .Cells(vFrm.ColExchangeRate.Index).Value = Me.objCD.DigitGrouping(Row.Ratio, vDefPrec)
                        Else
                            .Cells(vFrm.ColExchangeRate.Index).Value = Me.objCD.DigitGrouping(Row.Ratio, Core_Module.DefPrecExchangeRate, Core_Module.ClsCekData.EnumTipeRounding.Middlepoint)
                        End If
                        .Cells(vFrm.ColJob.Index).Value = Row.primary_job
                        .Cells(vFrm.ColRemarks.Index).Value = Row.Remarks
                    End With
                Next
            End If
            Dgv.AllowUserToAddRows = True
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    'Public Sub CetakJurnal_Cash_n_Bank1(ByVal vPrimary As Double, _
    'ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer)
    '    Dim ds As New DseTransaksiJournal
    '    Dim dta As New DseTransaksiJournalTableAdapters.DataTable2TableAdapter
    '    Dim tbl As DseTransaksiJournal.DataTable2DataTable = ds.DataTable2
    '    Dim row As DseTransaksiJournal.DataTable2Row
    '    Dim judul As String = ""
    '    Dim KetHeader As String = ""
    '    Dim KetItem As String = ""
    '    Dim InWords As String = ""
    '    Dim AmountHeader As Double = 0
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim FrmPilih As New frmPilih
    '    Dim rptCash As New CrpcCash
    '    Dim rptBank As New CrpcBank
    '    Dim Received_Payment As String = ""
    '    Try
    '        FrmPilih.ShowDialog()
    '        dta.Connection.ConnectionString = Core_Module.GetConstring
    '        dta.Fill(tbl, vPrimary)
    '        ' Dim baris As Integer = 10

    '        AmountHeader = 0
    '        For Each row In tbl.Rows
    '            AmountHeader = AmountHeader + row.Amount
    '            'If row.Currency_Code.Trim.ToUpper = "IDR" Then
    '            '    row.Remarks = row.Remarks
    '            'Else
    '            '    row.Remarks = row.Remarks & " - (" & Format(row.Amount_original, "###,###,##0.00") & "@" & Format(row.exchange_rate, "###,###,##0") & ")"
    '            'End If
    '        Next

    '        ' InWords = ObjCD.GetTerbilang(AmountHeader, Core_Module.ClsCekData.EnumStyleGetTerbilang.FirstLetterUpperCase)
    '        'AmountHeader = ObjCD.FormatHargaINA(CStr(AmountHeader))
    '        'Dim strCurrency As String = ""
    '        'If String.IsNullOrEmpty(DTACurr.ScalarQuery_GetCurrencyName(Core_Module.PrimaryBaseCurrency)) = False Then
    '        '    strCurrency = DTACurr.ScalarQuery_GetCurrencyName(Core_Module.PrimaryBaseCurrency).ToString
    '        'End If
    '        'KodeCur = ObjCur.GetCurrencyCode(Core_Module.PrimaryBaseCurrency)

    '        Dim strReceive As String = ""

    '        If FrmPilih.rboBank.Checked = True Then
    '            If tbl.Item(0).Type_InOut = ClsTransaksiJournal.EnumTipeInOut._In Then
    '                judul = "BANK RECEIPT VOUCHER"
    '                KetHeader = "(Debet)"
    '                KetItem = "(Credit)"
    '                Received_Payment = "RECEIVED FROM : "
    '                strReceive = "RECEIVED FOR"
    '            ElseIf tbl.Item(0).Type_InOut = ClsTransaksiJournal.EnumTipeInOut.out Then
    '                judul = "BANK PAYMENT VOUCHER"
    '                KetHeader = "(Credit)"
    '                KetItem = "(Debet)"
    '                Received_Payment = "PAYMENT TO : "
    '                strReceive = "PAYMENT FOR"
    '            End If
    '            Dim intBTMU As Integer = 0
    '            Dim intBCA As Integer = 0
    '            Dim intMIZUHO As Integer = 0
    '            Dim intMandiri As Integer = 0
    '            If FrmPilih.chkBTMU.Checked = True Then
    '                intBTMU = 1
    '            Else
    '                intBTMU = 0
    '            End If

    '            If FrmPilih.chkBCA.Checked = True Then
    '                intBCA = 1
    '            Else
    '                intBCA = 0
    '            End If

    '            If FrmPilih.chkMizuho.Checked = True Then
    '                intMIZUHO = 1
    '            Else
    '                intMIZUHO = 0
    '            End If

    '            If FrmPilih.chkMandiri.Checked = True Then
    '                intMandiri = 1
    '            Else
    '                intMandiri = 0
    '            End If

    '            rptBank.SetDataSource(ds)
    '            rptBank.SetParameterValue(rptBank.Parameter_Judul.ParameterFieldName, judul)
    '            rptBank.SetParameterValue(rptBank.Parameter_In_Word.ParameterFieldName, InWords)
    '            rptBank.SetParameterValue(rptBank.Parameter_KetHeader.ParameterFieldName, KetHeader)
    '            rptBank.SetParameterValue(rptBank.Parameter_KetItem.ParameterFieldName, KetItem)
    '            rptBank.SetParameterValue(rptBank.Parameter_Amount_Header.ParameterFieldName, AmountHeader)
    '            rptBank.SetParameterValue(rptBank.Parameter_Receive_Payment.ParameterFieldName, Received_Payment)
    '            rptBank.SetParameterValue(rptBank.Parameter_BCA.ParameterFieldName, intBCA)
    '            rptBank.SetParameterValue(rptBank.Parameter_BTMU.ParameterFieldName, intBTMU)
    '            rptBank.SetParameterValue(rptBank.Parameter_MANDIRI.ParameterFieldName, intMandiri)
    '            rptBank.SetParameterValue(rptBank.Parameter_MIZUHO.ParameterFieldName, intMIZUHO)
    '            rptBank.SetParameterValue(rptBank.Parameter_Receive_Payment_For.ParameterFieldName, strReceive)
    '            Crv.ReportSource = rptBank
    '        ElseIf FrmPilih.rboCash.Checked = True Then
    '            If tbl.Item(0).Type_InOut = ClsTransaksiJournal.EnumTipeInOut._In Then
    '                judul = "CASH RECEIPT VOUCHER"
    '                KetHeader = "(Debet)"
    '                KetItem = "(Credit)"
    '                Received_Payment = "Received from"
    '            ElseIf tbl.Item(0).Type_InOut = ClsTransaksiJournal.EnumTipeInOut.out Then
    '                judul = "CASH PAYMENT VOUCHER"
    '                KetHeader = "(Credit)"
    '                KetItem = "(Debet)"
    '                Received_Payment = "Payment to"
    '            End If
    '            rptCash.SetDataSource(ds)
    '            rptCash.SetParameterValue(rptCash.Parameter_Judul.ParameterFieldName, judul)
    '            rptCash.SetParameterValue(rptCash.Parameter_In_Word.ParameterFieldName, InWords)
    '            rptCash.SetParameterValue(rptCash.Parameter_KetHeader.ParameterFieldName, KetHeader)
    '            rptCash.SetParameterValue(rptCash.Parameter_KetItem.ParameterFieldName, KetItem)
    '            rptCash.SetParameterValue(rptCash.Parameter_Amount_Header.ParameterFieldName, AmountHeader)
    '            rptCash.SetParameterValue(rptCash.Parameter_Receive_Payment.ParameterFieldName, Received_Payment)
    '            Crv.ReportSource = rptCash
    '        End If

    '        Crv.Show()
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ds.Dispose()
    '        dta.Dispose()
    '        ObjCD = Nothing
    '    End Try
    'End Sub

    Public Function CetakCDNote(ByVal VPrimaryJournal As Double, ByVal Crv As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
                                    ByVal TipeTransaksi As EnumTipeTransaksi, ByVal TipeKartu As EnumTipeKartu, _
                                    ByVal vPrimaryCustSupp As Double, ByVal vPrimaryCurrAkunAR As Double, _
                                    ByVal vPrimaryCurrAkunReff As Double) As Boolean

        Dim ds As New DseTransaksiJournal
        Dim dta As New DseTransaksiJournalTableAdapters.tblCetakJurnalTableAdapter
        Dim tbl As DseTransaksiJournal.tblCetakJurnalDataTable = ds.tblCetakJurnal
        Dim report As New CrpCetakJurnalCDNote
        Dim report2 As New CrpCetakJurnalCDNote2
        Dim judul As String = ""
        Dim Label As String = ""
        Dim ObjCust As New Core_Module.clsMasterCustomer
        Dim ObjSupp As New Core_Module.clsMasterSupplier
        Dim ObjCurr As New Core_Module.clsMasterCurrency

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            dta.Fill(tbl, VPrimaryJournal)
            Dim BaseDeffPrec As Decimal = 0
            Dim strCustSupp As String = ""
            Dim BasePrimary As Double = -1

            BasePrimary = ObjCurr.GetPrimary("IDR")
            BaseDeffPrec = ObjCurr.GetDecimal_DefaultCurrency

            If TipeTransaksi = EnumTipeTransaksi.CREDIT_NOTE Then
                If TipeKartu = EnumTipeKartu.Customer Then
                    judul = "Customer's Credit Note"
                    Label = "Customer "
                    strCustSupp = ObjCust.GetCustomerCode(vPrimaryCustSupp) & " - " & ObjCust.GetCustName(vPrimaryCustSupp)
                Else
                    judul = "Supplier's Credit Note"
                    Label = "Supplier "
                    strCustSupp = ObjSupp.GetSupplierCode(vPrimaryCustSupp) & " - " & ObjSupp.GetSupplierName(vPrimaryCustSupp)
                End If
            Else
                If TipeKartu = EnumTipeKartu.Customer Then
                    judul = "Customer's Debit Note"
                    Label = "Customer "
                    strCustSupp = ObjCust.GetCustomerCode(vPrimaryCustSupp) & " - " & ObjCust.GetCustName(vPrimaryCustSupp)
                Else
                    judul = "Supplier's Debit Note"
                    Label = "Supplier "
                    strCustSupp = ObjSupp.GetSupplierCode(vPrimaryCustSupp) & " - " & ObjSupp.GetSupplierName(vPrimaryCustSupp)
                End If
            End If


            'If FrmPilih.RboA4.Checked = True Then
            If vPrimaryCurrAkunAR <> BasePrimary OrElse vPrimaryCurrAkunReff <> BasePrimary Then
                report2.SetDataSource(ds)
                report2.SetParameterValue(report2.Parameter_Judul.ParameterFieldName, judul)
                report2.SetParameterValue(report2.Parameter_LabelCustSupp.ParameterFieldName, Label)
                report2.SetParameterValue(report2.Parameter_CodeCustSupp.ParameterFieldName, strCustSupp)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, BaseDeffPrec)
                report2.SetParameterValue(report2.Parameter_Negatif.ParameterFieldName, BaseDeffPrec)
                Crv.ReportSource = report2
            Else
                report.SetDataSource(ds)
                report.SetParameterValue(report.Parameter_Judul.ParameterFieldName, judul)
                report.SetParameterValue(report.Parameter_LabelCustSupp.ParameterFieldName, Label)
                report.SetParameterValue(report.Parameter_CodeCustSupp.ParameterFieldName, strCustSupp)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, BaseDeffPrec)
                report.SetParameterValue(report.Parameter_Negatif.ParameterFieldName, BaseDeffPrec)
                Crv.ReportSource = report
            End If
            'ElseIf FrmPilih.RboA5.Checked = True Thens
            '    rptA5.SetDataSource(ds)
            '    rptA5.SetParameterValue(report.Parameter_Judul.ParameterFieldName, judul)
            '    Crv.ReportSource = rptA5
            'End If

            Crv.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function ApakahNoJurnalAda(ByVal vNo As String) As Boolean
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Jml As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Jml = DTA.ScalarQuery_GetJmlByNoJurnal(vNo)
            If Jml.HasValue = True Then
                If Jml.Value > 0 Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoJurnal(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiCreditNoteTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetNoJurnal(vPrimary)
            If String.IsNullOrEmpty(Retval) = True Then
                Return ""
            Else
                Return Retval
            End If
            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCustomer(ByVal vPrimaryTransaksi As Double, ByVal vTipeKartu As Integer) As Double
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_credit_noteTableAdapter
        Dim Retval As New Nullable(Of Double)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Retval = DTA.ScalarQuery_GetPrimaryCustomer(vPrimaryTransaksi, vTipeKartu)
            If Retval.HasValue = True Then
                If Retval.Value <> -1 Then
                    Return Retval.Value
                End If
            End If
            Return -1
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAdaNoJurnal(ByVal vNoJurnal As String) As Boolean
        Dim DTA As New DseTransaksiCreditNote_TransferJournalTableAdapters.tbltrans_jurnalTableAdapter
        Dim Count As New Nullable(Of Long)
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Count = DTA.ScalarQuery_GetJmlNoJur(vNoJurnal)
            If Count.HasValue = True Then
                If Count.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If
            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
