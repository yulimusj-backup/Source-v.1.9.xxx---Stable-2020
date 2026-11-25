Option Explicit On
Option Strict Off

Public Class ClsTransaksiQuotation
    Dim ObjCD As New Core_Module.ClsCekData
    Dim ObjUser As New Core_Module.ClsUser
    Dim ObjCompany As New Core_Module.ClsMasterCompanyData

    Private mPrimary As Double = -1
    ''' <summary>
    ''' digunakan untuk menggenerate no transaksi secara otomatis, default value = 0
    ''' </summary>
    ''' <remarks></remarks>
    Private mNoUrut As Double = 0
    Private mNoTransaksi As String = ""
    Private mTanggalTransaksi As Date = Now
    Private mPrimaryCust As Double = -1
    Private mIdentifier_CustAddress As String = ""
    Private mCustContactPerson As String = ""
    Private mCustPhone As String = ""
    Private mCustNPWP As String = ""
    Private mCustInquiryNo As String = ""
    Private mTermCondition As String = ""
    Private mPaymentTerm As String
    Private mPaymentDue As Integer = 0
    Private mReferencePaymentDate As EnumRefPaymentDate = EnumRefPaymentDate.InvoiceDate
    Private mRemarks As String = ""
    Private mPrimaryCurrency As Double = 1
    Private mSubTotal As Double = 0
    Private mUseTax As Integer = 0
    Private mTaxPercentage As Double = 0
    Private mTaxAmount As Double = 0
    Private mTotal As Double = 0
    Private mSalesStaffAuthorizationPrimary As Double = -1
    Private mSalesStaffAuthorizationDate As Date = Now
    Private mSalesManagerAuthorizationPrimary As Double = -1
    Private mSalesManagerAuthorizationDate As Date = Now
    Private mGeneralManagerAuthorizationPrimary As Double = -1
    Private mGeneralManagerAuthorizationDate As Date = Now
    Private mAuthorizationLevel As Integer = 0
    Private mAuthorizationLevelNeeded As Integer = 0
    Private mNamaSigned As String = ""
    Private mJabatanSIgned As String = ""
    Private mNeedDP As Integer = 0
    Private mDPPercent As Decimal = 0
    Private mDPAmount As Decimal = 0
    Private mDPFromPercent As Integer = 1
    Private mIdentifierAddressShipTo As String = ""
    Private mPrimaryCustomerShipTo As Double = -1
    Private mIdentifierAddressCustomerShipTo As String = ""

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumStatusQuo
        AllQuo
        AuthorizeLvl_1
        AuthorizeLvl_2
        AuthorizeLvl_3
        UnauthorizedOnly
    End Enum

    Public Property NeedDP() As Integer
        Get
            Return Me.mNeedDP
        End Get
        Set(ByVal value As Integer)
            Me.mNeedDP = value
        End Set
    End Property

    Public Property DPAmount() As Decimal
        Get
            Return Me.mDPAmount
        End Get
        Set(ByVal value As Decimal)
            Me.mDPAmount = value
        End Set
    End Property

    Public Property DPPercent() As Decimal
        Get
            Return Me.mDPPercent
        End Get
        Set(ByVal value As Decimal)
            mDPPercent = value
        End Set
    End Property

    Public Property DPFromPercent() As Decimal
        Get
            Return Me.mDPFromPercent
        End Get
        Set(ByVal value As Decimal)
            mDPFromPercent = value
        End Set
    End Property

    Public Property NamaSigned() As String
        Get
            Return Me.mNamaSigned
        End Get
        Set(ByVal value As String)
            mNamaSigned = value
        End Set
    End Property

    Public Property JabatanSigned() As String
        Get
            Return Me.mJabatanSIgned
        End Get
        Set(ByVal value As String)
            mJabatanSIgned = value
        End Set
    End Property

    Public Enum EnumRefPaymentDate As Integer
        InvoiceDate
        PODate
    End Enum

    Public Property AuthorizationLevelNeeded() As Integer
        Get
            Return Me.mAuthorizationLevelNeeded
        End Get
        Set(ByVal value As Integer)
            mAuthorizationLevelNeeded = value
        End Set
    End Property

    Public Property AuthorizationLevel() As Integer
        Get
            Return Me.mAuthorizationLevel
        End Get
        Set(ByVal value As Integer)
            mAuthorizationLevel = value
        End Set
    End Property

    Public Property SalesStaffAuthorizationPrimary() As Double
        Get
            Return Me.mSalesStaffAuthorizationPrimary
        End Get
        Set(ByVal value As Double)
            mSalesStaffAuthorizationPrimary = value
        End Set
    End Property

    Public Property SalesStaffAuthorizationDate() As Date
        Get
            Return Me.mSalesStaffAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mSalesStaffAuthorizationDate = value
        End Set
    End Property

    Public Property SalesManagerAuthorizationPrimary() As Double
        Get
            Return Me.mSalesManagerAuthorizationPrimary
        End Get
        Set(ByVal value As Double)
            mSalesManagerAuthorizationPrimary = value
        End Set
    End Property

    Public Property SalesManagerAuthorizationDate() As Date
        Get
            Return Me.mSalesManagerAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mSalesManagerAuthorizationDate = value
        End Set
    End Property

    Public Property GeneralManagerAuthorizationPrimary() As Double
        Get
            Return Me.mGeneralManagerAuthorizationPrimary
        End Get
        Set(ByVal value As Double)
            mGeneralManagerAuthorizationPrimary = value
        End Set
    End Property

    Public Property GeneralManagerAuthorizationDate() As Date
        Get
            Return Me.mGeneralManagerAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mGeneralManagerAuthorizationDate = value
        End Set
    End Property

    Public Property CustInquiryNo() As String
        Get
            Return Me.mCustInquiryNo
        End Get
        Set(ByVal value As String)
            mCustInquiryNo = value
        End Set
    End Property

    Public Property Total() As Double
        Get
            Return Me.mTotal
        End Get
        Set(ByVal value As Double)
            mTotal = value
        End Set
    End Property

    Public Property TaxAmount() As Double
        Get
            Return Me.mTaxAmount
        End Get
        Set(ByVal value As Double)
            mTaxAmount = value
        End Set
    End Property

    Public Property TaxPercentage() As Double
        Get
            Return Me.mTaxPercentage
        End Get
        Set(ByVal value As Double)
            mTaxPercentage = value
        End Set
    End Property

    Public Property UseTax() As Boolean
        Get
            If Me.mUseTax = 0 Then
                Return False
            Else
                Return True
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = False Then
                Me.mUseTax = 0
            Else
                Me.mUseTax = 1
            End If
        End Set
    End Property

    Public Property SubTotal() As Double
        Get
            Return Me.mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
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

    Public Property RefPaymentDate() As EnumRefPaymentDate
        Get
            Return Me.mReferencePaymentDate
        End Get
        Set(ByVal value As EnumRefPaymentDate)
            mReferencePaymentDate = value
        End Set
    End Property

    Public Property PaymentDue() As Integer
        Get
            Return Me.mPaymentDue
        End Get
        Set(ByVal value As Integer)
            mPaymentDue = value
        End Set
    End Property

    Public Property PaymentTerm() As String
        Get
            Return Me.mPaymentTerm
        End Get
        Set(ByVal value As String)
            mPaymentTerm = value
        End Set
    End Property

    Public Property TermCondition() As String
        Get
            Return Me.mTermCondition
        End Get
        Set(ByVal value As String)
            mTermCondition = value
        End Set
    End Property

    Public Property CustNPWP() As String
        Get
            Return Me.mCustNPWP
        End Get
        Set(ByVal value As String)
            mCustNPWP = value
        End Set
    End Property

    Public Property CustPhone() As String
        Get
            Return Me.mCustPhone
        End Get
        Set(ByVal value As String)
            mCustPhone = value
        End Set
    End Property

    Public Property CustContactPerson() As String
        Get
            Return Me.mCustContactPerson
        End Get
        Set(ByVal value As String)
            mCustContactPerson = value
        End Set
    End Property

    Public Property CustIdentifierAdd() As String
        Get
            Return Me.mIdentifier_CustAddress
        End Get
        Set(ByVal value As String)
            mIdentifier_CustAddress = value
        End Set
    End Property

    Public Property PrimaryCustomer() As Double
        Get
            Return Me.mPrimaryCust
        End Get
        Set(ByVal value As Double)
            mPrimaryCust = value
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
    Public Property IdentifierAddressShipTo() As String
        Get
            Return mIdentifierAddressShipTo
        End Get
        Set(ByVal value As String)
            mIdentifierAddressShipTo = value
        End Set
    End Property

    Public Property PrimaryCustomerShipTo() As Double
        Get
            Return mPrimaryCustomerShipTo
        End Get
        Set(ByVal value As Double)
            mPrimaryCustomerShipTo = value
        End Set
    End Property

    Public Property IdentifierAddressCustomerShipTo() As String
        Get
            Return mIdentifierAddressCustomerShipTo
        End Get
        Set(ByVal value As String)
            mIdentifierAddressCustomerShipTo = value
        End Set
    End Property


    Public Function ApakahSudahOtorisasi(ByVal vPrimaryQuotation As Double) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim TBL As New DseTransaksiQuotation.tbltrans_quotationDataTable
        Dim row As DseTransaksiQuotation.tbltrans_quotationRow

        Try
            If DTA.FillBy_Primary(TBL, vPrimaryQuotation) > 0 Then
                row = TBL.Rows(0)
                If row.Authorization_Level = row.Authorization_Level_Needed Then
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
            TBL.Dispose()
        End Try
    End Function

    Public Function GetNoBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetNoUrutMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetNoUrutMax(TanggalEnd, TanggalStart).Value + 1)
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
            Return NoUrut + "/" + "QUO" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            ObjCD = Nothing
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If DTA.ScalarQuery_GetJml(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJml(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
    ByVal vDgvItem As DataGridView, ByVal vDT_ItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiQuotation.DataTable2DataTable
        Dim row As DseTransaksiQuotation.DataTable2Row
        Dim ObjItem As New ClsTransaksiQuotationItem
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            'fil basis data
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mAuthorizationLevel = .Authorization_Level
                    Me.mAuthorizationLevelNeeded = .Authorization_Level_Needed
                    Me.mCustContactPerson = .Contact_Person
                    Me.mCustInquiryNo = .Cust_Inquiry_No
                    Me.mCustNPWP = .NPWP_No
                    Me.mCustPhone = .Cust_Phone_No
                    Me.mGeneralManagerAuthorizationDate = .General_Manager_Authorization_Date
                    Me.mGeneralManagerAuthorizationPrimary = .General_Manager_Authorization_Primary
                    Me.mIdentifier_CustAddress = .Indentifier_Alamat_Customer
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPaymentDue = .Payment_Due
                    Me.mPaymentTerm = .Payment_Term
                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mReferencePaymentDate = .Reference_Payment_Due
                    Me.mSalesManagerAuthorizationDate = .Sales_Manager_Authorization_Date
                    Me.mSalesManagerAuthorizationPrimary = .Sales_Manager_Authorization_Primary
                    Me.mSalesStaffAuthorizationDate = .Sales_Staff_Authorization_Date
                    Me.mSalesStaffAuthorizationPrimary = .Sales_Staff_Authorization_Primary
                    Me.mSubTotal = .Sub_Total
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxPercentage = .Tax
                    Me.mTermCondition = .Terms_N_Condition
                    Me.mTimestamp = .Timestamp
                    Me.mTotal = .Total
                    Me.mUseTax = .Use_Tax
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mNamaSigned = .Nama_Signed
                    Me.mJabatanSIgned = .Jabatan_Signed
                    Me.mNeedDP = .need_dp
                    Me.mDPAmount = .dp_amount
                    Me.mDPFromPercent = .dp_from_percent
                    Me.mDPPercent = .dp_percent
                    Me.mIdentifierAddressShipTo = .identifier_address_ship_to
                    Me.mPrimaryCustomerShipTo = .primary_customer_ship_to
                    Me.mIdentifierAddressCustomerShipTo = .identifier_address_customer_ship_to

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, vPrimaryTransaksi, vDT_ItemDelivery, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
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
            ObjCur = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
    Optional ByVal vTampilkanPesan As Boolean = True, Optional ByVal vUserCode As String = "", _
    Optional ByVal vSalesStrict As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            'jika bukan sales staff yang bersangkutan dan strict sales staff maka tidak boleh edit
            If vSalesStrict = True Then
                If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)) = True Then
                    If vTampilkanPesan = True Then
                        MsgBox("This transaction can be editted by: ", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                Else
                    If DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi) <> vUserCode Then
                        If vTampilkanPesan = True Then
                            MsgBox("This transaction can be editted by: " + DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi), _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'karena tidak ada transaksi yang merujuk quotation selalu boleh diedit selama 
            'belum diotorisasi
            If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryTransaksi).HasValue = False Then
            Else
                If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryTransaksi).Value <> 0 Then
                    If vTampilkanPesan = True Then
                        MsgBox("You cannot edit this data, cancel transaction authorization first.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'karena tidak ada transaksi yang merujuk quotation selalu boleh dihapus selama 
            'belum diotorisasi
            If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryTransaksi).HasValue = False Then
                Return True
            Else
                If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryTransaksi).Value <> 0 Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
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

            '2. Jika semua transaksi sales sedang dikunci

            '3. Jika sudah melewati batas penguncian transaksi

            '4. jika tidak ada item yang akan disimpan

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal i As FrmTransaksiQuotation, _
     ByVal vDT_ItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim DTA1 As New DseTransaksiQuotationTableAdapters.tbltrans_quotation3TableAdapter
        Dim ObjItem As New ClsTransaksiQuotationItem
        Dim rowDGV As DataGridViewRow
        Dim Dummy As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting

        Try
            If Me.ApakahDataBenar(True) = True Then
                ''hitung level otorisasi yang dibutuhkan
                Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.SalesQuotation_UseLvl2_Quotation)
                If Dummy = "" Then
                    Me.mAuthorizationLevelNeeded = 3
                Else
                    Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.SalesQuotation_UseLvl2_Limit)
                    If Dummy = "" Then
                        Me.mAuthorizationLevelNeeded = 3
                    Else
                        If Me.mTotal <= CDec(Dummy) Then
                            Me.mAuthorizationLevelNeeded = 2
                        Else
                            Me.mAuthorizationLevelNeeded = 3
                        End If
                    End If
                End If



                'insert query
                If DTA1.Insert(Me.PrimaryCustomer, Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), _
                Me.mTanggalTransaksi, Me.mIdentifier_CustAddress, Me.mCustContactPerson, Me.mCustPhone, _
                Me.mCustInquiryNo, Me.mPaymentTerm, Me.mPaymentDue, Me.mReferencePaymentDate, Me.mPrimaryCurrency, _
                Me.mSubTotal, Me.mUseTax, Me.mTaxPercentage, Me.mTaxAmount, Me.mTotal, Me.mTermCondition, _
                Me.mRemarks, -1, #1/1/1950#, -1, #1/1/1950#, -1, #1/1/1950#, Me.mAuthorizationLevelNeeded, 0, vUserCode, _
                Now, vUserCode, Now, Now, Me.mNamaSigned, Me.mJabatanSIgned, Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent, Me.mIdentifierAddressShipTo, _
                Me.mPrimaryCustomerShipTo, Me.mIdentifierAddressCustomerShipTo) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = CInt(Me.mNoTransaksi.Substring(0, 6))
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'sebelum simpan baru, update data identifier ada semua baris
                    For Each rowDGV In vdgvItem.Rows
                        rowDGV.Cells(i.ColIdentifier.Index).Value = Me.mPrimary.ToString + "-" + _
                        rowDGV.Cells(i.ColIdentifier.Index).Value.ToString.Substring((rowDGV.Cells(i.ColIdentifier.Index).Value.ToString.IndexOf("-", 1) + 1))
                    Next
                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, i, vDT_ItemDelivery)
                    'cek berkaitan nomor sama untuk beda bulan
                    Me.CekNomor(DTA)

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode, i, vDT_ItemDelivery)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjItem = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter)
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

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal i As FrmTransaksiQuotation, _
    ByVal vDT_ItemDelivery As DseTransaksiQuotation.tbltrans_quotation_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim DTA1 As New DseTransaksiQuotationTableAdapters.tbltrans_quotation3TableAdapter

        Dim ObjItem As New ClsTransaksiQuotationItem
        Dim Dummy As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting

        Try
            If Me.ApakahDataBenar(False) = True Then
                ''hitung level otorisasi yang dibutuhkan
                Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.SalesQuotation_UseLvl2_Quotation)
                If Dummy = "" Then
                    Me.mAuthorizationLevelNeeded = 3
                Else
                    Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, _
                    Core_Module.ClsTransaksiSetting.EnumNamaParameter.SalesQuotation_UseLvl2_Limit)
                    If Dummy = "" Then
                        Me.mAuthorizationLevelNeeded = 3
                    Else
                        If Me.mTotal <= CDec(Dummy) Then
                            Me.mAuthorizationLevelNeeded = 2
                        Else
                            Me.mAuthorizationLevelNeeded = 3
                        End If
                    End If
                End If

                If DTA1.Update(Me.mPrimaryCust, Me.mNoTransaksi, CInt(Me.mNoTransaksi.Substring(0, 6)), Me.mTanggalTransaksi, _
                Me.mIdentifier_CustAddress, Me.mCustContactPerson, Me.mCustPhone, Me.mCustInquiryNo, _
                Me.mPaymentTerm, Me.mPaymentDue, Me.mReferencePaymentDate, Me.mPrimaryCurrency, Me.mSubTotal, _
                Me.mUseTax, Me.mTaxPercentage, Me.mTaxAmount, Me.mTotal, Me.mTermCondition, Me.mRemarks, _
                Me.mAuthorizationLevelNeeded, Me.mAuthorizationLevel, vUserCode, Now, Now, _
                Me.mNamaSigned, Me.mJabatanSIgned, Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent, Me.mIdentifierAddressShipTo, _
                Me.mPrimaryCustomerShipTo, Me.mIdentifierAddressCustomerShipTo, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Me.mNoTransaksi.Substring(0, 6)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, i, vDT_ItemDelivery)
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
                Return Me.SimpanPerubahan(dgvItem, vUserCode, i, vDT_ItemDelivery)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            ObjItem = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi) = True Then
                'hapus data header, data item otomatis dihapuskan oleh server
                If DTA.Delete(vPrimaryTransaksi) > 0 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CetakQuotation(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Boolean
        'aul
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            'Dim ijk As New FrmCetakQuotation1
            Dim objCustom As New CustomSMESS.ClsTransaksiQuotation
            objCustom.CetakQuotation(CRV, vprimaryQuotaion)
            objCustom = Nothing
            'ijk.ShowDialog()
            'ijk.Dispose()
            Return False
        End If
        ObjSetting = Nothing
        '---

        'Dim report As New CrpCetakTransaksiQuotation
        'Dim DS As New DseTransaksiQuotation
        'Dim FormatDay As Integer = 1
        'Dim FormatMonth As Integer = 3
        'Dim FormatYears As Integer = 1
        'Dim FormatDecimal As Integer = 2
        Dim ij As New FrmCetakQuotation

        Try
            ij.Primary = vprimaryQuotaion
            'CetakQuotationVersiIndoPendek(CRV , vprimaryQuotaion)
            ij.ShowDialog()
            'Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            'DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            'DTAHeader.Dispose()
            'Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            'DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            'DTAIsi.Dispose()

            'report.SetDataSource(DS)
            'report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            'report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            'report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            'report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            'CRV.ReportSource = report
            'CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            'DS.Dispose()
            ij.Dispose()
        End Try
    End Function

    Public Function CetakQuotation_Batch(ByVal vListPrimaryQuotaion As List(Of Double), ByVal vLanguage As Integer, ByVal vVersion As Integer, _
    Optional ByVal vShowForm As Boolean = False, Optional ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer = Nothing, _
    Optional ByVal CRV2 As CrystalDecisions.Windows.Forms.CrystalReportViewer = Nothing) As Boolean
        'aul
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SalesQuotation, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            'Dim ijk As New FrmCetakQuotation1
            Dim objCustom As New CustomSMESS.ClsTransaksiQuotation
            objCustom.CetakQuotation_Batch(vListPrimaryQuotaion, vLanguage, vVersion, True)
            objCustom = Nothing
            'ijk.Dispose()
            Return False
        End If
        ObjSetting = Nothing
        '---

        Dim ij As New FrmCetakQuotation

        Dim DS As New DseTransaksiQuotation
        Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
        Dim TblHeaderDumm As New DseTransaksiQuotation.TblCetakHeaderDataTable
        Dim TblHeader As DseTransaksiQuotation.TblCetakHeaderDataTable = DS.TblCetakHeader
        Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
        Dim TblIsiDumm As New DseTransaksiQuotation.TblCetakIsiDataTable
        Dim TblIsi As DseTransaksiQuotation.TblCetakIsiDataTable = DS.TblCetakIsi
        Dim rowHeader, rowHeaderDumm As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi, rowIsiDumm As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim PrimaryMin As Double = -1
        Dim PrimaryMax As Double = -1
        Dim CompanyCity As String = ""

        Try
            If vShowForm = True Then
                ij.ListPrimary = vListPrimaryQuotaion
                ij.Batch = True
                ij.ShowDialog()
                Return True
            End If

            CompanyCity = Me.ObjCompany.GetCompany_City
            If CompanyCity <> "" Then
                CompanyCity = CompanyCity + ", "
            End If

            vListPrimaryQuotaion.Sort()
            PrimaryMin = vListPrimaryQuotaion.Item(0)
            PrimaryMax = vListPrimaryQuotaion.Item(vListPrimaryQuotaion.Count - 1)

            DTAHeader.FillBy_PrimaryQuotationMinMax(TblHeaderDumm, PrimaryMin, PrimaryMax)
            DTAIsi.FillBy_PrimaryQuotationMinMax(TblIsiDumm, PrimaryMin, PrimaryMax)

            Dim Filter As String = ""
            Dim Filter2 As String = ""
            Filter = " [Primary] = [Primary] "
            Filter2 = " [Primary Transaksi] = [Primary Transaksi] "

            If vListPrimaryQuotaion.Count > 0 Then
                Dim str As String = ""
                Dim str2 As String = ""
                For intA As Integer = 0 To vListPrimaryQuotaion.Count - 1
                    If str = "" Then
                        str &= " [Primary] = " & vListPrimaryQuotaion.Item(intA)
                    Else
                        str &= " OR [Primary] = " & vListPrimaryQuotaion.Item(intA)
                    End If

                    If str2 = "" Then
                        str2 &= " [Primary Transaksi] = " & vListPrimaryQuotaion.Item(intA)
                    Else
                        str2 &= " OR [Primary Transaksi] = " & vListPrimaryQuotaion.Item(intA)
                    End If
                Next
                Filter &= "And (" & str & ")"

                Filter2 &= "And (" & str2 & ")"
            End If

            TblHeaderDumm.DefaultView.RowFilter = Nothing
            TblHeaderDumm.DefaultView.RowFilter = Filter
            TblHeaderDumm.DefaultView.Sort = "[Primary]"
            While TblHeaderDumm.DefaultView.Count > 0
                rowHeaderDumm = TblHeaderDumm.DefaultView(0).Row
                rowHeader = TblHeader.NewTblCetakHeaderRow
                For ijk As Integer = 0 To TblHeaderDumm.Columns.Count - 1
                    rowHeader.Item(ijk) = rowHeaderDumm.Item(ijk)
                Next
                TblHeader.Rows.Add(rowHeader)
                TblHeaderDumm.Rows.Remove(rowHeaderDumm)
            End While

            TblIsiDumm.DefaultView.RowFilter = Nothing
            TblIsiDumm.DefaultView.RowFilter = Filter2
            TblIsiDumm.DefaultView.Sort = "[Primary Transaksi]"
            While TblIsiDumm.DefaultView.Count > 0
                rowIsiDumm = TblIsiDumm.DefaultView(0).Row
                rowIsi = TblIsi.NewTblCetakIsiRow
                For ijk As Integer = 0 To TblIsiDumm.Columns.Count - 1
                    rowIsi.Item(ijk) = rowIsiDumm.Item(ijk)
                Next
                TblIsi.Rows.Add(rowIsi)
                TblIsiDumm.Rows.Remove(rowIsiDumm)
            End While


            'rowHeader = DS.TblCetakHeader.Rows(0)
            'FormatDecimal = rowHeader.Default_Precision
            For Each rowIsi In TblIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next

            If vLanguage = 0 And vVersion = 0 Then
                Dim report As New CrpCetakTransaksiQuotation2_Batch 'inggris pendek
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, "")
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)
                report.SetParameterValue(report.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                CRV.ReportSource = report
                CRV.Show()
            ElseIf vLanguage = 0 And vVersion = 1 Then
                Dim report As New CrpCetakTransaksiQuotationPanjang2_Batch 'inggris panjang
                Dim report2 As New CrpCetakTransaksiQuotationDetail_Batch
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, "")
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)
                report.SetParameterValue(report.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                report2.SetDataSource(DS)
                report2.SetParameterValue(report2.Parameter_Currency.ParameterFieldName, "")
                report2.SetParameterValue(report2.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report2.SetParameterValue(report2.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report2.SetParameterValue(report2.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report2.SetParameterValue(report2.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                CRV.ReportSource = report
                CRV2.ReportSource = report2
                CRV.Show()
                CRV2.Show()
            ElseIf vLanguage = 1 And vVersion = 0 Then
                Dim report As New CrpCetakTransaksiQuotation_Batch 'indo pendek
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, "")
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)
                report.SetParameterValue(report.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                CRV.ReportSource = report
                CRV.Show()
            ElseIf vLanguage = 1 And vVersion = 1 Then
                Dim report As New CrpCetakTransaksiQuotationPanjang_Batch 'indo panjang
                Dim report2 As New CrpCetakTransaksiQuotationDetail_Batch
                report.SetDataSource(DS)
                report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, "")
                report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)
                report.SetParameterValue(report.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                report2.SetDataSource(DS)
                report2.SetParameterValue(report2.Parameter_Currency.ParameterFieldName, "")
                report2.SetParameterValue(report2.Parameter_FormatDay.ParameterFieldName, FormatDay)
                report2.SetParameterValue(report2.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                report2.SetParameterValue(report2.Parameter_FormatYears.ParameterFieldName, FormatYears)
                report2.SetParameterValue(report2.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                report2.SetParameterValue(report2.Parameter_MaxPrimary.ParameterFieldName, PrimaryMax)

                CRV.ReportSource = report
                CRV2.ReportSource = report2
                CRV.Show()
                CRV2.Show()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ij.Dispose()
            DTAHeader.Dispose()
            DTAIsi.Dispose()
            TblHeader.Dispose()
            TblHeaderDumm.Dispose()
            TblIsi.Dispose()
            TblIsiDumm.Dispose()
        End Try
    End Function

    ''Cetak Versi Detail
    Public Function CetakQuotationVersiDetail(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimaryQuotaion As Double) As Double
        Dim report As New CrpCetakTransaksiQuotationDetail
        Dim DS As New DseTransaksiQuotation
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim rowHeader As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            rowHeader = DS.TblCetakHeader.Rows(0)
            FormatDecimal = rowHeader.Default_Precision

            DTAHeader.Dispose()

            Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            DTAIsi.Dispose()

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, rowHeader.Currency_Code)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    ''Cetak Versi Bahasa Indonesia Pendek
    Public Function CetakQuotationVersiIndoPendek(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryQuotaion As Double, ByVal CompanyCity As String) As Boolean
        Dim report As New CrpCetakTransaksiQuotation
        Dim DS As New DseTransaksiQuotation
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim rowHeader As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            rowHeader = DS.TblCetakHeader.Rows(0)
            FormatDecimal = rowHeader.Default_Precision

            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            'While DS.TblCetakIsi.Rows.Count < 15
            'DS.TblCetakIsi.Rows.Add(vprimaryQuotaion, "-", "", "", 0, 0, 0, "")
            'End While

            DTAIsi.Dispose()

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, rowHeader.Currency_Code)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    ''Cetak Versi Bahasa English Pendek
    Public Function CetakQuotationVersiEnglishPendek(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryQuotaion As Double, ByVal CompanyCity As String) As Boolean
        Dim report As New CrpCetakTransaksiQuotation2
        Dim DS As New DseTransaksiQuotation
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim rowHeader As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            rowHeader = DS.TblCetakHeader.Rows(0)
            FormatDecimal = rowHeader.Default_Precision
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            'While DS.TblCetakIsi.Rows.Count < 15
            'DS.TblCetakIsi.Rows.Add(vprimaryQuotaion, "-", "", "", 0, 0, 0, "")
            'End While

            DTAIsi.Dispose()


            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, rowHeader.Currency_Code)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    ''Cetak Versi Bahasa English Panjang
    Public Function CetakQuotationVersiEnglishPanjang(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryQuotaion As Double, ByVal CompanyCity As String) As Boolean
        Dim report As New CrpCetakTransaksiQuotationPanjang2
        Dim DS As New DseTransaksiQuotation
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim rowHeader As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            rowHeader = DS.TblCetakHeader.Rows(0)
            FormatDecimal = rowHeader.Default_Precision
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            DTAIsi.Dispose()

            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, rowHeader.Currency_Code)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)

            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    ''Cetak Versi Bahasa Indonesia Panjang
    Public Function CetakQuotationVersiIndoPanjang(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, _
    ByVal vprimaryQuotaion As Double, ByVal CompanyCity As String) As Boolean
        Dim report As New CrpCetakTransaksiQuotationPanjang
        Dim DS As New DseTransaksiQuotation
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 3
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim rowHeader As DseTransaksiQuotation.TblCetakHeaderRow
        Dim rowIsi As DseTransaksiQuotation.TblCetakIsiRow
        Dim ObjProd As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiQuotationTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimaryQuotation(DS.TblCetakHeader, vprimaryQuotaion)
            DTAHeader.Dispose()
            rowHeader = DS.TblCetakHeader.Rows(0)
            FormatDecimal = rowHeader.Default_Precision
            Dim DTAIsi As New DseTransaksiQuotationTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimaryQuotation(DS.TblCetakIsi, vprimaryQuotaion)
            For Each rowIsi In DS.TblCetakIsi.Rows
                rowIsi.Decimal_Multi = ObjProd.GetAngkaDecimal(rowIsi.Primary_Product, rowIsi.Primary_Units_Multi)
            Next
            DTAIsi.Dispose()


            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_Currency.ParameterFieldName, rowHeader.Currency_Code)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            report.SetParameterValue(report.Parameter_Company_City.ParameterFieldName, CompanyCity)
            CRV.ReportSource = report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            ObjProd = Nothing
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimaryQuo As Double) As Date
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If DTA.ScalarQuery_GetTimestamp(vPrimaryQuo).HasValue = True Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimaryQuo).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal AuthorizedOnly As Boolean, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal vUserCode As String = "") As Long
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If AuthorizedOnly = False Then
                If vUserCode = "" Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        If DTA.ScalarQuery_GetJmlByNoTranskasiLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByNoTranskasiLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        If DTA.ScalarQuery_GetJmlByTanggalTransaksi(vTanggalSampai, vTanggalDari).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByTanggalTransaksi(vTanggalSampai, vTanggalDari).Value
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        If DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        If DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductDescLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        If DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductSpecLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        If DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustCodeLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        If DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        If DTA.ScalarQuery_GetJmlByCustInquiryLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustInquiryLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        If DTA.ScalarQuery_GetJmlByNoTransaksiLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE2(Criteria + "%", vUserCode)
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        If DTA.ScalarQuery_GetJmlByTanggalTransaksi2(vTanggalSampai, vTanggalDari, vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByTanggalTransaksi2(vTanggalSampai, vTanggalDari, vUserCode).Value
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        If DTA.ScalarQuery_GetJmlByProductCodeLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductCodeLIKE2(Criteria + "%", vUserCode).Value
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        If DTA.ScalarQuery_GetJmlByProductDescLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductDescLIKE2(Criteria + "%", vUserCode)
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        If DTA.ScalarQuery_GetJmlByProductSpecLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByProductSpecLIKE2(Criteria + "%", vUserCode)
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        If DTA.ScalarQuery_GetJmlByCustCodeLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustCodeLIKE2(Criteria + "%", vUserCode)
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        If DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%").HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustNameLIKE(Criteria + "%")
                        Else
                            Return 0
                        End If
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        If DTA.ScalarQuery_GetJmlByCustInquiryLIKE2(Criteria + "%", vUserCode).HasValue = True Then
                            Return DTA.ScalarQuery_GetJmlByCustInquiryLIKE2(Criteria + "%", vUserCode)
                        Else
                            Return 0
                        End If
                    End If
                End If
            Else
                If vUserCode = "" Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE_Authorized(Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        Return DTA.ScalarQuery_GetJmlByTanggalTransaksi_Authorized(vTanggalSampai, vTanggalDari).Value
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        Return DTA.ScalarQuery_GetJmlByProductCodeLIKE_Authorized(Criteria + "%").Value
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        Return DTA.ScalarQuery_GetJmlByProductDescLIKE_Authorized(Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        Return DTA.ScalarQuery_GetJmlByProductSpecLIKE_Authorized(Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        Return DTA.ScalarQuery_GetJmlByCustCodeLIKE_Authorized(Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        Return DTA.ScalarQuery_GetJmlByCustNameLIKE_Authorized(Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        Return DTA.ScalarQuery_GetJmlByCustInquiryLIKE_Authorized(Criteria + "%")
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE_Authorized2(Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        Return DTA.ScalarQuery_GetJmlByTanggalTransaksi_Authorized2(vTanggalSampai, vTanggalDari, vUserCode).Value
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        Return DTA.ScalarQuery_GetJmlByProductCodeLIKE_Authorized2(Criteria + "%", vUserCode).Value
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        Return DTA.ScalarQuery_GetJmlByProductDescLIKE_Authorized2(Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        Return DTA.ScalarQuery_GetJmlByProductSpecLIKE_Authorized2(Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        Return DTA.ScalarQuery_GetJmlByCustCodeLIKE_Authorized2(Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        Return DTA.ScalarQuery_GetJmlByCustNameLIKE_Authorized2(Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        Return DTA.ScalarQuery_GetJmlByCustInquiryLIKE_Authorized2(Criteria + "%", vUserCode)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiChecklistBoxOtorisasi(ByVal vLevel As Integer, _
    ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiQuotation.tbltrans_quotation1DataTable, _
    Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotation1TableAdapter

        Try
            'mengisikan semua quotation yang belum diotorisasi
            If vLevel = 0 Then
                If StrictSalesStaff = True Then
                    DTA.FillBy_UserNAuthoLevel(TBL, SalesStaff, vLevel)
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                End If
            ElseIf vLevel = 2 Then
                DTA.FillBy_MaxAuthorization(TBL, 2)
            Else
                DTA.FillBy_AuthorizationLevel(TBL, vLevel)
            End If

            'tampilkan pada checklist box
            clb.DataSource = TBL.DefaultView
            clb.DisplayMember = "No Transaksi"
            clb.Refresh()

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiChecklistBoxBatalOtorisasi(ByVal vLevel As Integer, _
    ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiQuotation.tbltrans_quotation1DataTable, _
    Optional ByVal StrictSalesStaff As Boolean = True, _
    Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotation1TableAdapter

        Try
            'mengisikan semua quotation yang belum diotorisasi
            If StrictSalesStaff = True Then
                If vLevel = 2 Then
                    DTA.FillBy_CancelAuthorizationSalesManager(TBL, 2, SalesStaff)
                ElseIf vLevel = 3 Then
                    DTA.FillBy_CancelGeneralManagerAuthorization(TBL, 3, SalesStaff)
                ElseIf vLevel = 1 Then
                    DTA.FillBy_UserNAuthoLevel(TBL, SalesStaff, vLevel)
                End If
            Else
                DTA.FillBy_AuthorizationLevel(TBL, vLevel)
            End If

            'tampilkan pada checklist box
            clb.DataSource = TBL.DefaultView
            clb.DisplayMember = "No Transaksi"
            clb.Refresh()

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
      ByVal DGV As DataGridView, ByVal TBL As DseTransaksiQuotation.DataTable3DataTable, _
      ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
      ByVal AuhtorizedOnly As Boolean, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
      Optional ByVal vTanggalSampai As Date = (#1/1/1950#), Optional ByVal vUserCode As String = "") As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.DataTable3TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            If AuhtorizedOnly = False Then
                If vUserCode = "" Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        QueryResult = DTA.FillBy_CustInquiryLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                        QueryResult = DTA.FillBy_CustCodeShipToLIKE(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                        QueryResult = DTA.FillBy_CustNameShipToLIKE(TBL, Criteria + "%")
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalSampaiDari2(TBL, vTanggalSampai, vTanggalDari, vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        QueryResult = DTA.FillBy_CustInquiryLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        QueryResult = DTA.FillBy_CustCodeLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                        QueryResult = DTA.FillBy_CustNameLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                        QueryResult = DTA.FillBy_CustCodeShiptoLIKE2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                        QueryResult = DTA.FillBy_CustNameShipToLIKE2(TBL, Criteria + "%", vUserCode)
                    End If
                End If
            Else
                If vUserCode = "" Then
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalSampaiDari_Authorized(TBL, vTanggalSampai, vTanggalDari)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        QueryResult = DTA.FillBy_CustInquiryLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        QueryResult = DTA.FillBy_CustCodeLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_CustNameLIKE_Authorized(TBL, Criteria + "%")
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                        QueryResult = DTA.FillBy_CustCodeShipToLIKE_Authorized(TBL, Criteria + "%")
                    End If
                Else
                    If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                        QueryResult = DTA.FillBy_NoTransaksiLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                        QueryResult = DTA.FillBy_TanggalSampaiDari_Authorized2(TBL, vTanggalSampai, vTanggalDari, vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerInquiryNo Then
                        QueryResult = DTA.FillBy_CustInquiryLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                        QueryResult = DTA.FillBy_CustCodeLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                        QueryResult = DTA.FillBy_ProductCodeLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                        QueryResult = DTA.FillBy_ProductDescLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                        QueryResult = DTA.FillBy_ProductSpecLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                        QueryResult = DTA.FillBy_CustCodeShiptoLIKE_Authorized2(TBL, Criteria + "%", vUserCode)
                    End If
                End If
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Transaction No"
                Dim ij As New DataGridViewCellStyle
                ij.Format = "dd MMM yyyy"
                DGV.Columns(2).DefaultCellStyle = ij
                DGV.Columns(2).HeaderText = "Date"
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mNoUrut = 0
            mNoTransaksi = ""
            mTanggalTransaksi = Now
            mPrimaryCust = -1
            mIdentifier_CustAddress = ""
            mCustContactPerson = ""
            mCustPhone = ""
            mCustNPWP = ""
            mCustInquiryNo = ""
            mTermCondition = ""
            mPaymentTerm = ""
            mPaymentDue = 0
            mReferencePaymentDate = EnumRefPaymentDate.InvoiceDate
            mRemarks = ""
            mPrimaryCurrency = 1
            mSubTotal = 0
            mUseTax = 0
            mTaxPercentage = 0
            mTaxAmount = 0
            mTotal = 0
            mSalesStaffAuthorizationPrimary = -1
            mSalesStaffAuthorizationDate = Now
            mSalesManagerAuthorizationPrimary = -1
            mSalesManagerAuthorizationDate = Now
            mGeneralManagerAuthorizationPrimary = -1
            mGeneralManagerAuthorizationDate = Now
            mAuthorizationLevel = 0
            mAuthorizationLevelNeeded = 0
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function CancelAuthorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
    ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim row As DseTransaksiQuotation.tbltrans_quotation1Row
        Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""
        Try
            'For Each rowview In clbAutho.CheckedItems
            '    row = rowview.Row
            '    If vAuthorizationLevel = 1 Then
            '        DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
            '        vAuthorizationLevel - 1, row.Primary)
            '    ElseIf vAuthorizationLevel = 2 Then
            '        DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
            '        vAuthorizationLevel - 1, row.Primary)
            '    ElseIf vAuthorizationLevel = 3 Then
            '        DTA.UpdateQuery_GeneralManagerAuthorization(-1, #1/1/1950#, _
            '        vAuthorizationLevel - 1, row.Primary)
            '    End If
            'Next

            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For i As Integer = 0 To clbAutho.CheckedItems.Count - 1
                rowview = clbAutho.CheckedItems(i)
                row = rowview.Row
                If vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                    'cancel otorisasi lanjutan jika memang punya hak akses atas hal tersebut, 
                    'dan yang melakukan otorisasi adalah orang yang bersangkutan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesStaff") = True Then
                        If Me.GetAuthorizeBy_Lvl1(row.Primary) = vPrimaryUser = -1 Then
                            DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                            vAuthorizationLevel - 1 - 1, row.Primary)
                        End If
                    End If
                ElseIf vAuthorizationLevel = 3 Then
                    DTA.UpdateQuery_GeneralManagerAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                    'cancel otorisasi lanjutan jika memang punya hak akses untuk melakukan hal tersebut
                    'hanya berlaku jika orang yang bersangkutan yang melakukan otorisasi
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesManager") = True Then
                        If Me.GetAuthorizeBy_Lvl2(row.Primary) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl2(row.Primary) = -1 Then
                            DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
                            vAuthorizationLevel - 1 - 1, row.Primary)

                            'lanjut jika yang mengotorisasi lvl 1 juga ybs dan ybs punya hak akses untuk cancel otorisasi
                            'pada lvl 1
                            If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesStaff") = True Then
                                If Me.GetAuthorizeBy_Lvl1(row.Primary) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl1(row.Primary) = -1 Then
                                    DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                                    vAuthorizationLevel - 1 - 1 - 1, row.Primary)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            'For Each rowview In clbAutho.CheckedItems

            'Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CancelAuthorize(ByVal vPrimaryQuo As Double, ByVal vPrimaryUser As String, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            'lv 3
            DTA.UpdateQuery_GeneralManagerAuthorization(-1, #1/1/1950#, _
                2, vPrimaryQuo)
            'lv 2
            DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
                1, vPrimaryQuo)
            'lv 1
            DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                0, vPrimaryQuo)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Authorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
    ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Dim row As DseTransaksiQuotation.tbltrans_quotation1Row
        Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""
        Try
            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For Each rowview In clbAutho.CheckedItems
                row = rowview.Row
                If vAuthorizationLevel = 0 Then
                    DTA.UpdateQuery_SalesStaffAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                    'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesManager") = True Then
                        DTA.UpdateQuery_SalesManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1 + 1, row.Primary)
                        'otorisasi lanjutan jika punya hak otorisasi
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiGeneralManager") = True Then
                            DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                            vAuthorizationLevel + 1 + 1 + 1, row.Primary)
                        End If
                    End If
                ElseIf vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_SalesManagerAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                    'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiGeneralManager") = True Then
                        DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1 + 1, row.Primary)
                    End If
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Authorize(ByVal vPrimaryQuo As Double, ByVal vPrimaryUser As String, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            'lvl 1
            DTA.UpdateQuery_SalesStaffAuthorization(vPrimaryUser, Now, _
                1, vPrimaryQuo)
            'lvl 2
            DTA.UpdateQuery_SalesManagerAuthorization(vPrimaryUser, Now, _
                2, vPrimaryQuo)
            'lvl 3
            DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
            3, vPrimaryQuo)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    'AUL

    ''' <summary>
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada Quotation dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryQuo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl1(ByVal vPrimaryQuo As Double) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl1(vPrimaryQuo).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl1(vPrimaryQuo).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada Quotation dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryQuo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl2(ByVal vPrimaryQuo As Double) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl2(vPrimaryQuo).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl2(vPrimaryQuo).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada Quotation dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryQuo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl3(ByVal vPrimaryQuo As Double) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl3(vPrimaryQuo).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl3(vPrimaryQuo).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRatioDefault(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Decimal
        Dim dta As New DseTransaksiQuotationTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim ratio_ As Nullable(Of Decimal) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            ratio_ = dta.ScalarQuery_GetRatioDefault(vPrimaryProduct, vPrimaryUnit)
            If ratio_.HasValue = True Then
                Return ratio_.Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    Public Function GetRatioReference(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Decimal
        Dim dta As New DseTransaksiQuotationTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim ratio_ As Nullable(Of Decimal) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            ratio_ = dta.ScalarQuery_GetRatioReferrence(vPrimaryProduct, vPrimaryUnit)
            If ratio_.HasValue = True Then
                Return ratio_.Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        End Try
    End Function

    'aul
    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vNoTrans).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vNoTrans).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiQuo(ByVal vprimary As Double) As String
        Dim dta As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(dta.ScalarQuery_GetNoTransaksi(vprimary)) = False Then
                Return dta.ScalarQuery_GetNoTransaksi(vprimary)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetAuthorizationLevel(ByVal vPrimaryQuo As Double) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryQuo).HasValue Then
                Return DTA.ScalarQuery_GetAuthorizationLevel(vPrimaryQuo).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAuthorizationLevelNeeded(ByVal vPrimaryQuo As Double) As Double
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetAuthorizationLevelNeeded(vPrimaryQuo).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgvBoxAlreadyAuthorize(ByVal vLevel As Integer, _
    ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiQuotation.tbltrans_qoutation_authDataTable, _
    ByVal vDgvUser As DataGridView, _
    Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
    Optional ByVal vQuoNo As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
    Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vStatusQuo As Integer = EnumStatusQuo.AllQuo, _
    Optional ByVal vQuoDate As Boolean = False) As Boolean
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_qoutation_authTableAdapter
        Dim row As DseTransaksiQuotation.tbltrans_qoutation_authRow
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView
        Dim StatusAuth As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'If vPoDate = False Then
            '    If vStatusQuo = EnumStatusQuo.AllPO Then
            '        DTA.Fill(TBL)
            '    ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_1 Then
            '        DTA.FillBy_AuthorizationLevel(TBL, 1)
            '    ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_2 Then
            '        DTA.FillBy_AuthorizationLevel(TBL, 2)
            '    ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_3 Then
            '        DTA.FillBy_AuthorizationLevelDirector(TBL, 3)
            '    ElseIf vStatusQuo = EnumStatusQuo.UnauthorizedOnly Then
            '        DTA.FillBy_AuthorizationLevel(TBL, 0)
            '    End If
            'Else
            If vStatusQuo = EnumStatusQuo.AllQuo Then
                DTA.Fill_TglDariSampai(TBL, vDateFrom, vDateTo)
            ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_1 Then
                DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, 1, vDateFrom, vDateTo)
            ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_2 Then
                DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, 2, vDateFrom, vDateTo)
            ElseIf vStatusQuo = EnumStatusQuo.AuthorizeLvl_3 Then
                DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, 3, vDateFrom, vDateTo)
            ElseIf vStatusQuo = EnumStatusQuo.UnauthorizedOnly Then
                DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, 0, vDateFrom, vDateTo)
            End If
            'End If

            If vQuoNo <> "" Then
                filter = "[No Transaksi] = '" & vQuoNo.ToString & "'"
            End If
            If vDgvUser.Rows.Count > 1 Then
                For Each r As DataGridViewRow In vDgvUser.Rows
                    With r
                        If String.IsNullOrEmpty(.Cells(vFrm.ColCreatedBy.Index).Value) = False Then
                            If String.IsNullOrEmpty(filter2) = True Then
                                filter2 = filter2 & "[Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            Else
                                filter2 = filter2 & " or [Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            End If
                        End If
                    End With
                Next
            End If
            If vQuoNo <> "" And filter2 <> "" Then
                filter = filter & " AND (" & filter2 & ")"
            ElseIf filter = "" Then
                filter = filter2
            End If

            'tampilkan pada datagridview box
            If Dgv.ColumnCount > 0 Then
                Dgv.Rows.Clear()
                TBL.DefaultView.RowFilter = Nothing
                TBL.DefaultView.RowFilter = filter
                For Each dv In TBL.DefaultView
                    row = dv.Row
                    Dgv.Rows.Add()
                    rowDGV = Dgv.Rows(Dgv.RowCount - 1)
                    rowDGV.Cells(vFrm.ColAA_Primary.Index).Value = row.Primary
                    rowDGV.Cells(vFrm.ColAA_QuoNo.Index).Value = row.No_Transaksi
                    If row.Authorization_Level = 1 Then
                        StatusAuth = "Authorized By Sales Staff"
                    ElseIf row.Authorization_Level = 2 Then
                        StatusAuth = "Authorized By Sales Manager"
                    ElseIf row.Authorization_Level = 3 Then
                        StatusAuth = "Authorized By Director"
                    Else
                        StatusAuth = "-"
                    End If
                    rowDGV.Cells(vFrm.ColAA_Status.Index).Value = StatusAuth
                    rowDGV.Cells(vFrm.ColAA_CustomerName.Index).Value = row.Customer_Name
                Next
            End If
            Dgv.Refresh()

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgvBoxOtorisasi(ByVal vLevel As Integer, _
   ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiQuotation.tbltrans_qoutation_authDataTable, _
   ByVal vDgvUser As DataGridView, _
   Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
   Optional ByVal vQuoNo As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
   Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vQuoDate As Boolean = False) As Boolean
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_qoutation_authTableAdapter
        Dim row As DseTransaksiQuotation.tbltrans_qoutation_authRow
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None



            'mengisikan semua quotation yang belum diotorisasi
            If vQuoDate = False Then
                If vLevel = 0 Then
                    If StrictSalesStaff = True Then
                        DTA.FillBy_UserNAuthorizationLevel(TBL, SalesStaff, vLevel)
                    Else
                        DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                    End If
                ElseIf vLevel = 2 Then
                    DTA.FillBy_MaxAuthorization(TBL, 2)
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                End If
            Else
                'filter dengan tanggal
                If vLevel = 0 Then
                    If StrictSalesStaff = True Then
                        DTA.FillBy_UserNAuthorizationLevelTglDariSampai(TBL, SalesStaff, vLevel, vDateFrom, vDateTo)
                    Else
                        DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, vLevel, vDateFrom, vDateTo)
                    End If
                ElseIf vLevel = 2 Then
                    DTA.FillBy_MaxAuthorizationTglDariSampai(TBL, 2, vDateFrom, vDateTo)
                Else
                    DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, vLevel, vDateFrom, vDateTo)
                End If
            End If

            If vQuoNo <> "" Then
                filter = "[No Transaksi] = '" & vQuoNo.ToString & "'"
            End If
            If vDgvUser.Rows.Count > 1 Then
                For Each r As DataGridViewRow In vDgvUser.Rows
                    With r
                        If String.IsNullOrEmpty(.Cells(vFrm.ColCreatedBy.Index).Value) = False Then
                            If String.IsNullOrEmpty(filter2) = True Then
                                filter2 = filter2 & "[Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            Else
                                filter2 = filter2 & " or [Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            End If

                        End If
                    End With
                Next
            End If

            If vQuoNo <> "" And filter2 <> "" Then
                filter = filter & " AND (" & filter2 & ")"
            ElseIf filter = "" Then
                filter = filter2
            End If

            'tampilkan pada datagridview box
            If Dgv.ColumnCount > 0 Then
                Dgv.Rows.Clear()
                TBL.DefaultView.RowFilter = Nothing
                TBL.DefaultView.RowFilter = filter
                For Each dv In TBL.DefaultView
                    row = dv.Row
                    Dgv.Rows.Add()
                    rowDGV = Dgv.Rows(Dgv.RowCount - 1)
                    rowDGV.Cells(vFrm.ColDgvAuth_Primary.Index).Value = row.Primary
                    rowDGV.Cells(vFrm.ColDgvAuth_QuoNo.Index).Value = row.No_Transaksi
                    rowDGV.Cells(vFrm.ColDgvAuth_CustomerName.Index).Value = row.Customer_Name
                Next
            End If
            Dgv.Refresh()

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDgvBoxBatalOtorisasi(ByVal vLevel As Integer, _
    ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiQuotation.tbltrans_qoutation_authDataTable, _
    ByVal vDgvUser As DataGridView, _
    Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
    Optional ByVal vQuoNo As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
    Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vQuoDate As Boolean = False) As Boolean
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_qoutation_authTableAdapter
        Dim row As DseTransaksiQuotation.tbltrans_qoutation_authRow
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'mengisikan semua quotation yang belum diotorisasi
            If vQuoDate = False Then
                If StrictSalesStaff = True Then
                    If vLevel = 2 Then
                        DTA.FillBy_CancelAuthorizationSalesManager(TBL, SalesStaff, vLevel)
                    ElseIf vLevel = 3 Then
                        DTA.FillBy_CancelAuthorizationGeneralManager(TBL, SalesStaff, vLevel)
                    ElseIf vLevel = 1 Then
                        DTA.FillBy_UserNAuthorizationLevel(TBL, SalesStaff, vLevel)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                End If
            Else
                'filter dengan tanggal
                If StrictSalesStaff = True Then
                    If vLevel = 2 Then
                        DTA.FillBy_CancelAuthorizationSalesManagerTglDariSampai(TBL, 2, vDateFrom, vDateTo, SalesStaff)
                    ElseIf vLevel = 3 Then
                        DTA.FillBy_CancelAuthorizationGeneralManagerTglDariSampai(TBL, 3, vDateFrom, vDateTo, SalesStaff)
                    ElseIf vLevel = 1 Then
                        DTA.FillBy_UserNAuthorizationLevelTglDariSampai(TBL, SalesStaff, vLevel, vDateFrom, vDateTo)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevelTglDariSampai(TBL, vLevel, vDateFrom, vDateTo)
                End If
            End If

            If vQuoNo <> "" Then
                filter = "[No Transaksi] = '" & vQuoNo.ToString & "'"
            End If
            If vDgvUser.Rows.Count > 1 Then
                For Each r As DataGridViewRow In vDgvUser.Rows
                    With r
                        If String.IsNullOrEmpty(.Cells(vFrm.ColCreatedBy.Index).Value) = False Then
                            If String.IsNullOrEmpty(filter2) = True Then
                                filter2 = filter2 & "[Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            Else
                                filter2 = filter2 & " or [Created By] = '" & .Cells(vFrm.ColCreatedBy.Index).Value & "'"
                            End If

                        End If
                    End With
                Next
            End If

            If vQuoNo <> "" And filter2 <> "" Then
                filter = filter & " AND (" & filter2 & ")"
            ElseIf filter = "" Then
                filter = filter2
            End If

            'tampilkan pada datagridview box
            If Dgv.ColumnCount > 0 Then
                Dgv.Rows.Clear()
                TBL.DefaultView.RowFilter = Nothing
                TBL.DefaultView.RowFilter = filter
                For Each dv In TBL.DefaultView
                    row = dv.Row
                    Dgv.Rows.Add()
                    rowDGV = Dgv.Rows(Dgv.RowCount - 1)
                    rowDGV.Cells(vFrm.ColDgvUnAuth_Primary.Index).Value = row.Primary
                    rowDGV.Cells(vFrm.ColDgvUnAuth_QuoNo.Index).Value = row.No_Transaksi
                    rowDGV.Cells(vFrm.ColDgvUnAuth_CustomerName.Index).Value = row.Customer_Name
                Next
            End If
            Dgv.Refresh()

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

            'sampe sini return true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SaveUserItem(ByVal vUserCode As String, ByVal vDgvUser As DataGridView) As Boolean
        Dim dta As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_item_userTableAdapter
        Dim tbl As New DseTransaksiQuotation.tbltrans_quotation_item_userDataTable
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim rowDgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_UserCode(tbl, vUserCode) > 0 Then
                dta.Delete(vUserCode)
            End If
            For i As Integer = 0 To vDgvUser.Rows.Count - 2
                rowDgv = vDgvUser.Rows(i)
                dta.Insert(vUserCode, rowDgv.Cells(vFrm.ColCreatedBy.Index).Value)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    Public Function LoadUserItem(ByVal vUserCode As String, ByVal vDgvUser As DataGridView) As Boolean
        Dim dta As New DseTransaksiQuotationTableAdapters.tbltrans_quotation_item_userTableAdapter
        Dim tbl As New DseTransaksiQuotation.tbltrans_quotation_item_userDataTable
        Dim row As DseTransaksiQuotation.tbltrans_quotation_item_userRow
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim rowDgv As DataGridViewRow
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.FillBy_UserCode(tbl, vUserCode) > 0 Then
                vDgvUser.Rows.Clear()
                For Each row In tbl
                    vDgvUser.Rows.Add()
                    rowDgv = vDgvUser.Rows(vDgvUser.RowCount - 1)
                    rowDgv.Cells(vFrm.ColCreatedBy.Index).Value = row.Ref_User
                    rowDgv.Cells(vFrm.ColCreatedBy_Name.Index).Value = ObjUser.GetNama(row.Ref_User)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
            tbl.Dispose()
        End Try
    End Function

    '============= Taufik ========================='
    Public Function authorizeQuotation(ByVal Dgv As DataGridView, ByVal vPrimaryUser As String, _
          ByVal vAuthorizationLevel As Integer) As Boolean
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        'Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        'Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""
        Dim Primary_ As Double = -1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For Each rowDgv As DataGridViewRow In Dgv.Rows
                If rowDgv.Cells(vFrm.ColDgvAuth_Authorize.Index).Value = Boolean.TrueString Then
                    Primary_ = rowDgv.Cells(vFrm.ColDgvAuth_Primary.Index).Value
                    If vAuthorizationLevel = 0 Then
                        DTA.UpdateQuery_SalesStaffAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1, Primary_)
                        'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesManager") = True Then
                            DTA.UpdateQuery_SalesManagerAuthorization(vPrimaryUser, Now, _
                            vAuthorizationLevel + 1 + 1, Primary_)
                            'otorisasi lanjutan jika punya hak otorisasi
                            If GetAuthorizationLevelNeeded(Primary_) = 3 Then
                                If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiGeneralManager") = True Then
                                    DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                                    vAuthorizationLevel + 1 + 1 + 1, Primary_)
                                End If
                            End If
                        End If
                    ElseIf vAuthorizationLevel = 1 Then
                        DTA.UpdateQuery_SalesManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1, Primary_)
                        'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiGeneralManager") = True Then
                            If GetAuthorizationLevelNeeded(Primary_) = 3 Then
                                DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                                vAuthorizationLevel + 1 + 1, Primary_)
                            End If
                        End If
                    ElseIf vAuthorizationLevel = 2 Then
                        DTA.UpdateQuery_GeneralManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1, Primary_)
                    End If
                End If
            Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Function CancelAuthorizeQuotation(ByVal Dgv As DataGridView, ByVal vPrimaryUser As String, _
   ByVal vAuthorizationLevel As Integer) As Boolean
        Dim vFrm As New FrmTransaksiQuotationOtorisasi2
        Dim DTA As New DseTransaksiQuotationTableAdapters.tbltrans_quotationTableAdapter
        'Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        'Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""
        Dim Primary_ As Double = -1
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For Each rowDgv As DataGridViewRow In Dgv.Rows
                If rowDgv.Cells(vFrm.ColDgvUnAuth_Cancel.Index).Value = Boolean.TrueString Then
                    Primary_ = rowDgv.Cells(vFrm.ColDgvUnAuth_Primary.Index).Value

                    If vAuthorizationLevel = 1 Then
                        DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                    ElseIf vAuthorizationLevel = 2 Then
                        DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                        'cancel otorisasi lanjutan jika memang punya hak akses atas hal tersebut, 
                        'dan yang melakukan otorisasi adalah orang yang bersangkutan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesStaff") = True Then
                            If Me.GetAuthorizeBy_Lvl1(Primary_) = vPrimaryUser = -1 Then
                                DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                                vAuthorizationLevel - 1 - 1, Primary_)
                            End If
                        End If
                    ElseIf vAuthorizationLevel = 3 Then
                        DTA.UpdateQuery_GeneralManagerAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                        'cancel otorisasi lanjutan jika memang punya hak akses untuk melakukan hal tersebut
                        'hanya berlaku jika orang yang bersangkutan yang melakukan otorisasi
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesManager") = True Then
                            If Me.GetAuthorizeBy_Lvl2(Primary_) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl2(Primary_) = -1 Then
                                DTA.UpdateQuery_SalesManagerAuthorization(-1, #1/1/1950#, _
                                vAuthorizationLevel - 1 - 1, Primary_)

                                'lanjut jika yang mengotorisasi lvl 1 juga ybs dan ybs punya hak akses untuk cancel otorisasi
                                'pada lvl 1
                                If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiQuotationOtorisasiSalesStaff") = True Then
                                    If Me.GetAuthorizeBy_Lvl1(Primary_) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl1(Primary_) = -1 Then
                                        DTA.UpdateQuery_SalesStaffAuthorization(-1, #1/1/1950#, _
                                        vAuthorizationLevel - 1 - 1 - 1, Primary_)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            'For Each rowview In clbAutho.CheckedItems

            'Next

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function
End Class
