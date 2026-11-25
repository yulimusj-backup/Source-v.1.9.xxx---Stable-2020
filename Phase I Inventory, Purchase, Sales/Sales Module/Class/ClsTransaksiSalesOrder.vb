Option Explicit On
Option Strict Off

Public Class ClsTransaksiSalesOrder
    Dim ObjCD As New Core_Module.ClsCekData

    'dalam 1 so tidak boleh ada 2 item dengan product code sama hanya berbeda pricing saja
    'dalam 1 so tidak boleh ada 2 item dengan product code sama hanya berbeda pricing saja
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
    Private mCustPONo As String = ""
    Private mCustPODate As Date = Now
    Private mPaymentTerm As String = ""
    Private mPaymentDue As Integer = 0
    Private mReferencePaymentDate As EnumRefPaymentDate = EnumRefPaymentDate.InvoiceDate
    Private mRemarks As String = ""
    Private mPrimaryCurrency As Double = 1
    Private mSubTotal As Double = 0
    Private mUseTax As Integer = 0
    Private mTaxPercentage As Double = 0
    Private mTaxAmount As Double = 0
    Private mTotal As Double = 0
    Private mShipVia As String = ""
    Private mOrderBy As EnumOrderBy = EnumOrderBy.PO
    Private mJenisFakturPajak As EnumJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
    Private mPriceCondition As EnumPriceCondition = EnumPriceCondition.Exclude
    Private mPriceConditonSama As Integer = 0
    Private mDiskonSama As Integer = 0
    Private mReferenceDiskon As Double = 0
    Private mExchangeRate As Decimal = 1
    Private mPrimarySalesPerson As Double = -1
    Private mIdentifierAddressShipTo As String = ""
    Private mPrimaryCustomerShipTo As Double = -1
    Private mIdentifierAddressCustomerShipTo As String = ""
    Private mNeedDP As Integer = 0
    Private mDPPercent As Decimal = 0
    Private mDPAmount As Decimal = 0
    Private mDPFromPercent As Integer = 1

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    Public Enum EnumPriceCondition As Integer
        Exclude
        Include
    End Enum

    Public Enum EnumRefPaymentDate As Integer
        InvoiceDate
        PODate
    End Enum

    Public Enum EnumOrderBy As Integer
        PO
        Telepon
        Fax
        Email
    End Enum

    Public Enum EnumJenisFakturPajak As Integer
        FakturPajakSederhana
        FakturPajakStandar
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

    Public Property ExchangeRate() As Decimal
        Get
            Return Me.mExchangeRate
        End Get
        Set(ByVal value As Decimal)
            mExchangeRate = value
        End Set
    End Property

    Public Property PrimarySalesPerson() As Double
        Get
            Return Me.mPrimarySalesPerson
        End Get
        Set(ByVal value As Double)
            mPrimarySalesPerson = value
        End Set
    End Property

    Public Property JenisFakturPajak() As EnumJenisFakturPajak
        Get
            Return Me.mJenisFakturPajak
        End Get
        Set(ByVal value As EnumJenisFakturPajak)
            mJenisFakturPajak = value
        End Set
    End Property

    Public Property ReferenceDiscount() As Double
        Get
            Return Me.mReferenceDiskon
        End Get
        Set(ByVal value As Double)
            mReferenceDiskon = value
        End Set
    End Property

    Public Property DiskonSama() As Boolean
        Get
            If Me.mDiskonSama = 1 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mDiskonSama = 1
            Else
                Me.mDiskonSama = 0
            End If
        End Set
    End Property

    Public Property PriceConditionSama() As Boolean
        Get
            If Me.mPriceConditonSama = 1 Then
                Return True
            Else
                Return False
            End If
        End Get
        Set(ByVal value As Boolean)
            If value = True Then
                Me.mPriceConditonSama = 1
            Else
                Me.mPriceConditonSama = 0
            End If
        End Set
    End Property

    Public Property PriceCondition() As EnumPriceCondition
        Get
            Return Me.mPriceCondition
        End Get
        Set(ByVal value As EnumPriceCondition)
            mPriceCondition = value
        End Set
    End Property

    Public Property OrderBy() As EnumOrderBy
        Get
            Return Me.mOrderBy
        End Get
        Set(ByVal value As EnumOrderBy)
            Me.mOrderBy = value
        End Set
    End Property

    Public Property ShipVia() As String
        Get
            Return Me.mShipVia
        End Get
        Set(ByVal value As String)
            mShipVia = value
        End Set
    End Property

    Public Property CustPONo() As String
        Get
            Return Me.mCustPONo
        End Get
        Set(ByVal value As String)
            Me.mCustPONo = value
        End Set
    End Property
    Public Property CustPODate() As Date
        Get
            Return Me.mCustPODate
        End Get
        Set(ByVal value As Date)
            mCustPODate = value
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

    Public Function GetCustPONo(ByVal vPrimary As Double, ByVal DTA As DseTransaksiSOTableAdapters.tbltrans_soTableAdapter) As String
        Try
            Return DTA.ScalarQuery_GetCustPONo(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetCustPONo(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Try
            Return DTA.ScalarQuery_GetCustPONo(vPrimary)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCustPONumber(ByVal vSONumber As String) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoPO(vSONumber)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoPO(vSONumber)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNOCustPOBaru(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Dummy As String
        Dim i As Integer = 0

        Try
            'tentukan tanggal awal dan akhir periode
            TanggalEnd = CDate("12/31/" + CStr(vTanggalTransaksi.Year))
            TanggalStart = CDate("1/1/" + CStr(vTanggalTransaksi.Year))

            ' get jumlah yang ada
            NoUrut = DTA.ScalarQuery_GetCustPONo_Max(TanggalEnd, TanggalStart)
            If NoUrut = "" Then
                'transaksi pertama pada tahun yang bersangkutan
                'generate no cust PO
                NoUrut = CStr(vTanggalTransaksi.Year) + "/001"
            Else
                'increment
                i = CInt(NoUrut.Substring(NoUrut.LastIndexOf("/") + 1))
                i += 1
                Dummy = CStr(i)
                While Dummy.Length < 3
                    Dummy = "0" + Dummy
                End While
                NoUrut = CStr(vTanggalTransaksi.Year) + "/" + Dummy
                'cek apakah sudah ada, bila sudah maka increment
                While DTA.ScalarQuery_GetJmlByCustPONo(NoUrut).Value > 0
                    i = CInt(NoUrut.Substring(NoUrut.LastIndexOf("/") + 1))
                    i += 1
                    Dummy = CStr(i)
                    While Dummy.Length < 3
                        Dummy = "0" + Dummy
                    End While
                    NoUrut = CStr(vTanggalTransaksi.Year) + "/" + Dummy
                End While
                'pastika terdiri dari 3 karakter
            End If

            'nilai balikan
            Return NoUrut
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetNoBaru_Obs(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
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
            Return NoUrut + "/" + "SO" + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            'cari tanggal sampai dan dari
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_GetUrutanMax(TanggalEnd, TanggalStart).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            Return Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, _
            Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function ApakahSudahTerkirimSemua(ByVal vPrimarySO As Double) As Boolean
        Dim DTA As New dseLaporanSalesOrderBalance2TableAdapters.laporan_sales_orderbalance_view9bTableAdapter

        Try
            Return DTA.ScalarQuery_GetTotalOB(vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

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
    ByVal vDgvItem As DataGridView, ByVal vDT_ItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_soDataTable
        Dim row As DseTransaksiSO.tbltrans_soRow
        Dim ObjItem As New ClsTransaksiSalesOrderItem
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            'fil basis data
            If DTA.FillBy_Primary(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mCustContactPerson = .Contact_Person
                    Me.mCustPhone = .Cust_Phone_No
                    Me.mCustPONo = .Cust_PO_No
                    Me.mIdentifier_CustAddress = .Indentifier_Alamat_Customer
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Urutan
                    Me.mPrimary = .Primary
                    Me.mPaymentDue = .Payment_Due
                    Me.mPaymentTerm = .Payment_Term
                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mPrimaryCust = .Primary_Customer
                    Me.mReferencePaymentDate = .Reference_Payment_Due
                    Me.mSubTotal = .Sub_Total
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTaxPercentage = .Tax
                    Me.mTotal = .Total
                    Me.mUseTax = .Use_Tax
                    Me.mRemarks = .Remarks
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.DiskonSama = .Discount_Sama
                    Me.mJenisFakturPajak = .Jenis_Faktur_Pajak
                    Me.mOrderBy = .Order_By
                    Me.mPriceCondition = .Price_Condition
                    Me.PriceConditionSama = .Price_Condtion_Sama
                    Me.mReferenceDiskon = .Reference_Discount
                    Me.mShipVia = .Ship_Via
                    Me.mPrimarySalesPerson = .primary_sales_person
                    Me.mExchangeRate = .Exchange_Rate
                    Me.mNeedDP = .need_dp
                    Me.mDPAmount = .dp_amount
                    Me.mDPFromPercent = .dp_from_percent
                    Me.mDPPercent = .dp_percent

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                    If .Cust_PO_Date = "#1/1/1950#" Then
                        Me.mCustPODate = Now
                    Else
                        Me.mCustPODate = .Cust_PO_Date
                    End If
                    Me.mIdentifierAddressShipTo = .identifier_address_ship_to
                    Me.mPrimaryCustomerShipTo = .primary_customer_ship_to
                    Me.mIdentifierAddressCustomerShipTo = .identifier_address_customer_ship_to

                End With
                'salin data item delivery
                ObjItem.IsiDGV(vDgvItem, Me.mPrimary, vDT_ItemDelivery, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency))
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
    ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim ObjUser As New Core_Module.ClsUser

        Try
            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit SO:
            '1. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
                If CreatedBy = vSalesStaff Then
                Else
                    If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSO_Override_Edit") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            '2. SO yang berasal dari quick invoice tidak boleh diedit.
            If DTA.ScalarQuery_GetPrimaryQuickInvoice(vPrimaryTransaksi).HasValue = True Then
                Dim PrimaryQuickInvoice As Double = DTA.ScalarQuery_GetPrimaryQuickInvoice(vPrimaryTransaksi).Value
                If PrimaryQuickInvoice <> -1 Then
                    If TampilkanPesan = True Then
                        Dim ObjInv As New ClsTransaksiSalesInvoice
                        MsgBox("Please use quick invoice form to edit this data for invoice: " + _
                        ObjInv.GetNoInvoice(PrimaryQuickInvoice), MsgBoxStyle.Information, "Information")
                        ObjInv = Nothing
                    End If
                    Return False
                End If
            End If
            'tidak ada batasan lain untuk edit transaksi karena, pada dasarnya batasan edit
            'sebenarnya adalah pada item, jadi kalo rajin mungkin bisa ditambahin:
            'JIKA SEMUA ITEM TIDAK BISA DIEDIT MAKA TRANSAKSI TIDAK BISA DIEDIT
            'msgbox("belum")

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim DTA1 As New DseTransaksiSalesInvoiceTableAdapters.tbltrans_salesinvoice5TableAdapter
        Dim TBL1 As New DseTransaksiSalesInvoice.tbltrans_salesinvoice5DataTable
        Dim row1 As DseTransaksiSalesInvoice.tbltrans_salesinvoice5Row
        Dim Dummy As String = ""
        Dim objUser As New Core_Module.ClsUser
        Dim DTA2 As New Core_Module.DseBalanceDPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TBL2 As New Core_Module.DseBalanceDP.tbltrans_historicalbalancemanual_dpDataTable
        Dim row2 As Core_Module.DseBalanceDP.tbltrans_historicalbalancemanual_dpRow
        Dim TipeBalanceDP As Core_Module.ClsTransaksiBalanceDP.EnumTipeDP = Core_Module.ClsTransaksiBalanceDP.EnumTipeDP.DP_Invoice

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString

            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'data SO boleh dihapus jika
            '1. tidak ada DO yang merujuk pada so ybs
            Dim ObjDOItem As New ClsTransaksiDOItem
            If ObjDOItem.GetJmlBySO(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this SO because there are DO refering to this SO", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjDOItem = Nothing
            '2. tidak ada SOC yang merujuk ke transaksi ybs
            Dim ObjSOC As New ClsTransaksiSOC
            If ObjSOC.GetJmlBySO(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this SO because there are SOC refering to this SO", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSOC = Nothing
            '3. tidak ada booking inv yang merujuk ke SO tersebut
            Dim ObjBookInv As New ClsTransaksiBookingInv
            If ObjBookInv.GetJmlBySO(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this SO because there are inventory booking data refering to this SO", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjBookInv = Nothing


            '5. Jika strict sales Staff delete maka yang boleh hapus hanya sales staff yang bersangkutan
            Dim CreatedBy As String = DTA.ScalarQuery_GetCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If objUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiSO_Override_Delete") = False Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            '6. tidak bisa dihapuskan bila ada down payment yang merujuk pada SO tersebut
            DTA1.Fill(TBL1, vPrimaryTransaksi)
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    If Dummy = "" Then
                        Dummy = row1.No_Transaksi
                    Else
                        Dummy = Dummy + ", " + row1.No_Transaksi
                    End If
                Next

                MsgBox("This sales order refered by down payment invoice: " + Dummy + ". You cannot delete this data.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'tidak boleh hapus jika sudah di refer oleh balance dp
            Dummy = ""
            DTA2.FillBy_PrimaryDocAcuan(TBL2, CInt(TipeBalanceDP), vPrimaryTransaksi)
            If TBL2.Rows.Count > 0 Then
                For Each row2 In TBL2.Rows
                    If Dummy = "" Then
                        Dummy = row2.DP_No
                    Else
                        Dummy = Dummy + ", " + row2.DP_No
                    End If
                Next

                MsgBox("This sales order refered by down payment invoice (begining balance): " + Dummy + ". You cannot delete this data.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            objUser = Nothing
            DTA2.Dispose()
            TBL2.Dispose()
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
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        MsgBox("Sales order number duplication, please change sales order number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        End If
                    End If
                Else
                    If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                        If Me.GetPrimary(Me.mNoTransaksi) <> Me.mPrimary Then
                            MsgBox("Sales order number duplication, please change sales order number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If
            '5. jika item yang disimpan mengakibatkan order balance minus

            '6. jika item yang disimpan nilainya lebih kecil dari yang sudah dibooking via IPO--> hati - hati ini

            '7. jika item yang disimpan nilainya lebih kecil dari yang sudah booking inv



            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub SetAsQuickSO(ByVal vPrimaryQuickInvoice As Double, ByVal vPrimarySO As Double)
        Dim DTASO As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTASO.UpdateQuery_SetPrimaryQuickInvoice(vPrimaryQuickInvoice, vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
        End Try
    End Sub

    Public Sub SetOrderBalance_QuickInvoice_Nol(ByVal vPrimarySO As Double)
        Dim DTASO As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_itemTableAdapter

        Try
            DTASO.UpdateQuery_SetSOBNol(vPrimarySO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTASO.Dispose()
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSO, _
    ByVal vDT_ItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable, Optional ByVal vCekDataBenar As Boolean = True, _
    Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim ObjItem As New ClsTransaksiSalesOrderItem
        Dim Urutan As Double = -1
        Dim rowDGV As DataGridViewRow

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            If Me.ApakahDataBenar(True) = True Then
                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                'hitung level otorisasi yang dibutuhkan
                'insert query
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, Me.mTanggalTransaksi)

                If DTA.Insert(Me.PrimaryCustomer, Me.mNoTransaksi, Urutan, _
                Me.mTanggalTransaksi, Me.mIdentifier_CustAddress, Me.mCustContactPerson, Me.mCustPhone, _
                Me.mPaymentTerm, Me.ShipVia, Me.mPaymentDue, Me.mReferencePaymentDate, Me.mPrimaryCurrency, _
                Me.mSubTotal, Me.mUseTax, Me.mTaxPercentage, Me.mTaxAmount, Me.mTotal, Me.OrderBy, _
                Me.CustPONo, Me.mJenisFakturPajak, Me.mPriceCondition, Me.mPriceConditonSama, Me.mDiskonSama, _
                Me.mReferenceDiskon, Me.mRemarks, CreatedBy_, CreatedDate_, ModifiedBy_, ModifiedDate_, TimeStamp_, _
                Me.mPrimarySalesPerson, Me.mExchangeRate, Me.mCustPODate, Me.mIdentifierAddressShipTo, _
                Me.mPrimaryCustomerShipTo, Me.mIdentifierAddressCustomerShipTo, Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mNoTransaksi).Value
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    If vCekDataBenar = True Then
                        'update data berkaitan dengan cust po no
                        If Me.mOrderBy <> EnumOrderBy.PO Then
                            While DTA.ScalarQuery_GetJmlByCustPONo(Me.mCustPONo).Value > 1
                                If DTA.ScalarQuery_GetPrimaryMaxByCustPONO(Me.mCustPONo).Value = _
                                Me.mPrimary Then
                                    'get no baru
                                    Me.mCustPONo = Me.GetNOCustPOBaru(Me.mTanggalTransaksi)
                                    'update no-nya
                                    DTA.UpdateQuery_CustPONO(Me.mCustPONo, Me.mPrimary)
                                End If
                            End While
                        End If
                    End If

                    'sebelum simpan baru, update data identifier ada semua baris
                    For Each rowDGV In vdgvItem.Rows
                        rowDGV.Cells(vFrm.ColIdentifier.Index).Value = Me.mPrimary.ToString + "-" + _
                        rowDGV.Cells(vFrm.ColIdentifier.Index).Value.ToString.Substring((rowDGV.Cells(vFrm.ColIdentifier.Index).Value.ToString.IndexOf("-", 1) + 1))
                    Next
                    'simpan data item baru
                    Me.UpdateSpareQty(Me.mPrimary, vFrm)
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vFrm, vDT_ItemDelivery)
                    If vCekDataBenar = True Then
                        'set order balance
                        ObjItem.Set_Order_Balance(Me.mPrimary)
                        'cek berkaitan nomor sama untuk beda bulan
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO) = True Then
                            Me.CekNomor(DTA, Me.ObjNumbering.GetResetType(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO))
                        End If
                    End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                Return Me.SimpanBaru(vdgvItem, vUserCode, vFrm, vDT_ItemDelivery, vCekDataBenar, vFromCutOff)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Private Sub CekNomor(ByVal DTA As DseTransaksiSOTableAdapters.tbltrans_soTableAdapter, ByVal vEnumTipeReset As Core_Module.ClsMasterTransactionNumbering.EnumReset)
        Dim TanggalDari, TanggalSampai As Date
        Dim Urutan As Integer = Me.mNoUrut
        Dim LoopGo As Boolean = True
        Dim JmlTrans As Integer
        Dim LoopCount As Integer
        Dim ObjCD As New Core_Module.ClsCekData
        Dim PrimaryMax As Double = -1

        Try
            If vEnumTipeReset = Core_Module.ClsMasterTransactionNumbering.EnumReset.Monthly Then
                TanggalDari = ObjCD.GetFirstDateofTheMonth(Me.mTanggalTransaksi)
            Else
                TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
            End If
            'TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
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
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                        Me.mNoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, Me.mTanggalTransaksi)
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

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiSO, _
    ByVal vDT_ItemDelivery As DseTransaksiSO.tbltrans_so_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim ObjItem As New ClsTransaksiSalesOrderItem
        Dim Urutan As Double = -1
        Dim ObjItemBookingINV As New ClsTransaksiBookingInvItem
        Try
            If Me.ApakahDataBenar(False) = True Then
                'hitung level otorisasi yang dibutuhkan
                Urutan = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, _
                Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.SO, Me.mTanggalTransaksi)

                If DTA.Update(Me.mPrimaryCust, Me.mNoTransaksi, Urutan, Me.mTanggalTransaksi, _
                Me.mIdentifier_CustAddress, Me.mCustContactPerson, Me.mCustPhone, _
                Me.mPaymentTerm, Me.mShipVia, Me.mPaymentDue, Me.mReferencePaymentDate, Me.mPrimaryCurrency, Me.mSubTotal, _
                Me.mUseTax, Me.mTaxPercentage, Me.mTaxAmount, Me.mTotal, Me.mOrderBy, Me.mCustPONo _
                , Me.mJenisFakturPajak, Me.mPriceCondition, Me.mPriceConditonSama, Me.mDiskonSama, Me.mReferenceDiskon, _
                Me.mRemarks, vUserCode, Now, Now, Me.mPrimarySalesPerson, Me.mExchangeRate, Me.mCustPODate, _
                Me.mIdentifierAddressShipTo, Me.mPrimaryCustomerShipTo, Me.mIdentifierAddressCustomerShipTo, _
                Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent, Me.mPrimary, Me.mTimestamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mNoUrut = Urutan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    Me.UpdateSpareQty(Me.mPrimary, vFrm)
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm, vDT_ItemDelivery)
                    'set order balance
                    ObjItem.Set_Order_Balance(Me.mPrimary)
                    '2026-11-07 IM, tambah update ke inv booking, jika ada item yang dihapus
                    ObjItemBookingINV.updateItemBooking(dgvItem, Me.mPrimary, vFrm)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi)
                Return Me.SimpanPerubahan(dgvItem, vUserCode, vFrm, vDT_ItemDelivery)
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjItemBookingINV = Nothing
        End Try
    End Function

    Private Sub UpdateSpareQty(ByVal vPrimaryTransaksi As Double, ByVal vFrmSO As FrmTransaksiSO)
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_so_item2TableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_so_item2DataTable
        Dim row As DseTransaksiSO.tbltrans_so_item2Row
        Dim rowDGV As DataGridViewRow

        Try
            'load data lama dari db
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryTransaksi)

            'update qty spare = 0
            For Each rowDGV In vFrmSO.dgvItem.Rows
                rowDGV.Cells(vFrmSO.ColQtySpare.Index).Value = 0

                'untuk setiap baris so, tentukan nilai qty spare lama dan copy ke dalam kolomnya
                For Each row In TBL.Rows
                    If row.Primary_Transaksi = vPrimaryTransaksi And row.Primary_Product = rowDGV.Cells(vFrmSO.ColPrimaryCodeNo.Index).Value Then
                        rowDGV.Cells(vFrmSO.ColQtySpare.Index).Value = row.MP_Spare_Qty

                        Exit For
                    End If
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            UserCode) = True Then
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

    Public Function GetTaxPersentage(ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            Return DTA.ScalarQuery_GetTaxPercentage(vPrimarySO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUseTax(ByVal vPrimarySO As Double) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetUseTax(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetUseTax(vPrimarySO).Value
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSubTotal(ByVal vPrimarySO As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetSubTotal(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetSubTotal(vPrimarySO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTanggal(ByVal vPrimarySO As Double) As Date
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggal(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggal(vPrimarySO).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimaryQuo As Double) As Date
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

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

    Public Function GetPaymentTerm(ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPaymentDuedate(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetPaymentDuedate(vPrimarySO).Value
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

    Public Function GetPrimaryBy_CustPONumber(ByVal vCust_PO_Number As String) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary_ByCustPONumber(vCust_PO_Number).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary_ByCustPONumber(vCust_PO_Number).Value
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

    Public Function GetPrimary(ByVal vNoSO As String) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimary(vNoSO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimary(vNoSO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimarySalesPerson(ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimarySalesPerson(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimarySalesPerson(vPrimarySO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimarySO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCustomer(ByVal vPrimarySO As Double) As Double
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If DTA.ScalarQuery_GetPrimaryCust(vPrimarySO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCust(vPrimarySO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoSO_ByPrimaryInvoiceQuick(ByVal vPrimary As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim RetVal As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTA.ScalarQuery_GetNoTransaksi_ByPrimaryInvoiceQuick(vPrimary)

            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return RetVal
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoSO_ByCustomerPONO(ByVal vCustomerPONO As String) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoTransaksi(vCustomerPONO)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoTransaksi(vCustomerPONO)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoSO(ByVal vPrimarySO As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSONo(vPrimarySO)) = False Then
                Return DTA.ScalarQuery_GetSONo(vPrimarySO)
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

    Public Function GetNoSO(ByVal vPrimarySO As Double, ByVal DTA As DseTransaksiSOTableAdapters.tbltrans_soTableAdapter) As String
        Try
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetSONo(vPrimarySO)) = False Then
                Return DTA.ScalarQuery_GetSONo(vPrimarySO)
            Else
                Return ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Sub GetCustCode_CustName_CustPONO_ContactPerson_PhoneNo(ByVal vPrimarySO As Double, ByRef CustName As String, _
    ByRef CustCode As String, ByRef CustPONO As String, ByRef ContactPerson As String, ByRef PhoneNo As String)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiSO.DataTable2DataTable
        Dim row As DseTransaksiSO.DataTable2Row

        Try
            If DTA.FillBy_PrimarySO(TBL, vPrimarySO) > 0 Then
                row = TBL.Rows(0)
                With row
                    CustName = .Customer_Name
                    CustCode = .Customer_Code
                    CustPONO = .Cust_Phone_No
                    ContactPerson = .Contact_Person
                    PhoneNo = .Cust_Phone_No
                End With
            Else
                CustName = ""
                CustCode = ""
                CustPONO = ""
                ContactPerson = ""
                PhoneNo = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub GetCustCode_CustName_CustPONO_SalesOrderNumber(ByVal vPrimarySO As Double, ByRef CustName As String, _
    ByRef CustCode As String, ByRef CustPONO As String, ByRef SO_Number As String, ByVal CustContact As String)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable2TableAdapter
        Dim TBL As New DseTransaksiSO.DataTable2DataTable
        Dim row As DseTransaksiSO.DataTable2Row

        Try
            If DTA.FillBy_PrimarySO(TBL, vPrimarySO) > 0 Then
                row = TBL.Rows(0)
                With row
                    CustCode = .Customer_Code
                    CustName = .Customer_Name
                    CustPONO = .Cust_PO_No
                    SO_Number = .SO_Number
                    CustContact = .Contact_Person
                End With
            Else
                CustName = ""
                CustCode = ""
                CustPONO = ""
                SO_Number = ""
                CustContact = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub GetPajakNJenisFakturNPrimaryCurrency(ByVal vPrimarySO As Double, _
    ByRef vPrimaryCurrency As Double, ByRef vJenisFaktur As EnumJenisFakturPajak, ByRef vPersenPajak _
    As Double)
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_soDataTable
        Dim row As DseTransaksiSO.tbltrans_soRow

        Try
            If DTA.FillBy_Primary(TBL, vPrimarySO) > 0 Then
                row = TBL.Rows(0)
                vPrimaryCurrency = row.Primary_Currency
                vPersenPajak = row.Tax
                vJenisFaktur = row.Jenis_Faktur_Pajak
            Else
                vPrimaryCurrency = -1
                vPersenPajak = 0
                vJenisFaktur = EnumJenisFakturPajak.FakturPajakSederhana
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Sub

    Public Sub GetAR_Detail_OverDue(ByVal vPrimaryCustomer As Double, ByVal vDueDate As Date, ByVal vFrmReminder As FrmTransaksiSO_ReminderAR)
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As DseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim i As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim rowDGV As DataGridViewRow

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryCustomer, vDueDate)
            For Each row In TBL.Rows
                If row.primary_currency <> 1 Then
                    i += 1
                End If

                If row.Tipe = CInt(TipeBalance_Inv) Then
                    row.Invoice_Number_Manual = ObjInv.GetNoInvoice(row.Primary_Invoice)
                ElseIf row.Tipe = CInt(TipeBalance_SR) Then
                    row.Invoice_Number_Manual = ObjSR.GetSRNO(row.Primary_Invoice)
                End If

                vFrmReminder.dgvItem.Rows.Add()
                rowDGV = vFrmReminder.dgvItem.Rows(vFrmReminder.dgvItem.Rows.Count - 1)
                rowDGV.Cells(vFrmReminder.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                rowDGV.Cells(vFrmReminder.ColBalance.Index).Value = ObjCD.DigitGrouping(row.Balance, DefPrec)
                rowDGV.Cells(vFrmReminder.ColCurrencyOriginal.Index).Value = row.Currency_Code
                rowDGV.Cells(vFrmReminder.ColDueDate.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                rowDGV.Cells(vFrmReminder.ColInvNo.Index).Value = row.Invoice_Number_Manual
                rowDGV.Cells(vFrmReminder.ColNo.Index).Value = vFrmReminder.dgvItem.Rows.Count
            Next

            vFrmReminder.ColBalance.HeaderText = "Balance (" + ObjCur.GetCurrencyCode(1) + ")"
            If i = 0 Then
                vFrmReminder.ColBalanceOriginal.Visible = False
                vFrmReminder.ColCurrencyOriginal.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjInv = Nothing
            ObjSR = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Public Function GetAR_Detail_OverDue(ByVal vPrimaryCustomer As Double, ByVal vDueDate As Date) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As DseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim Retval As String = ""
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim i As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryCustomer, vDueDate)
            For Each row In TBL.Rows
                i += 1
                If row.Tipe = CInt(TipeBalance_Inv) Then
                    row.Invoice_Number_Manual = ObjInv.GetNoInvoice(row.Primary_Invoice)
                ElseIf row.Tipe = CInt(TipeBalance_SR) Then
                    row.Invoice_Number_Manual = ObjSR.GetSRNO(row.Primary_Invoice)
                End If

                If Retval = "" Then
                    If row.primary_currency = 1 Then
                        Retval = i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    Else
                        Retval = i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + "(IDR " + ObjCD.DigitGrouping(row.Balance, DefPrec) + ")" + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    End If
                Else
                    If row.primary_currency = 1 Then
                        Retval = Retval + vbCrLf + i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    Else
                        Retval = Retval + vbCrLf + i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + "(IDR " + ObjCD.DigitGrouping(row.Balance, DefPrec) + ")" + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjInv = Nothing
            ObjSR = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Sub GetAR_Detail_ExceedsLimit(ByVal vPrimaryCustomer As Double, ByVal vFrmReminder As FrmTransaksiSO_ReminderAR)
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As DseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim i As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim rowDGV As DataGridViewRow

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_All(TBL, vPrimaryCustomer)

            For Each row In TBL.Rows
                If row.primary_currency <> 1 Then
                    i += 1
                End If
                If row.Tipe = CInt(TipeBalance_Inv) Then
                    row.Invoice_Number_Manual = ObjInv.GetNoInvoice(row.Primary_Invoice)
                ElseIf row.Tipe = CInt(TipeBalance_SR) Then
                    row.Invoice_Number_Manual = ObjSR.GetSRNO(row.Primary_Invoice)
                End If

                vFrmReminder.dgvItem.Rows.Add()
                rowDGV = vFrmReminder.dgvItem.Rows(vFrmReminder.dgvItem.Rows.Count - 1)
                rowDGV.Cells(vFrmReminder.ColBalanceOriginal.Index).Value = ObjCD.DigitGrouping(row.balance_original, row.Default_Precision)
                rowDGV.Cells(vFrmReminder.ColBalance.Index).Value = ObjCD.DigitGrouping(row.Balance, DefPrec)
                rowDGV.Cells(vFrmReminder.ColCurrencyOriginal.Index).Value = row.Currency_Code
                rowDGV.Cells(vFrmReminder.ColDueDate.Index).Value = Format(row.Due_Date, "dd MMM yyyy")
                rowDGV.Cells(vFrmReminder.ColInvNo.Index).Value = row.Invoice_Number_Manual
                rowDGV.Cells(vFrmReminder.ColNo.Index).Value = vFrmReminder.dgvItem.Rows.Count
            Next

            vFrmReminder.ColBalance.HeaderText = "Balance (" + ObjCur.GetCurrencyCode(1) + ")"
            If i = 0 Then
                vFrmReminder.ColBalanceOriginal.Visible = False
                vFrmReminder.ColCurrencyOriginal.Visible = False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjInv = Nothing
            ObjSR = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Sub

    Public Function GetAR_Detail_ExceedsLimit(ByVal vPrimaryCustomer As Double) As String
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As DseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim Retval As String = ""
        Dim ObjInv As New ClsTransaksiSalesInvoice
        Dim ObjSR As New ClsTransaksiSalesReturn
        Dim i As Integer = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DefPrec = ObjCur.GetDecimal_DefaultCurrency
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_All(TBL, vPrimaryCustomer)
            For Each row In TBL.Rows
                i += 1
                If row.Tipe = CInt(TipeBalance_Inv) Then
                    row.Invoice_Number_Manual = ObjInv.GetNoInvoice(row.Primary_Invoice)
                ElseIf row.Tipe = CInt(TipeBalance_SR) Then
                    row.Invoice_Number_Manual = ObjSR.GetSRNO(row.Primary_Invoice)
                End If
                If Retval = "" Then
                    If row.primary_currency = 1 Then
                        Retval = i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    Else
                        Retval = i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + "(IDR " + ObjCD.DigitGrouping(row.Balance, DefPrec) + ")" + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    End If
                Else
                    If row.primary_currency = 1 Then
                        Retval = Retval + vbCrLf + i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    Else
                        Retval = Retval + vbCrLf + i.ToString + ". " + row.Invoice_Number_Manual + " - Balance: " + _
                        row.Currency_Code + " " + ObjCD.DigitGrouping(row.balance_original, row.Default_Precision) + "(IDR " + ObjCD.DigitGrouping(row.Balance, DefPrec) + ")" + vbCrLf + _
                        " - Due Date: " + Format(row.Due_Date, "dd MMM yyyy")
                    End If
                End If
            Next

            Return Retval
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjInv = Nothing
            ObjSR = Nothing
            ObjCD = Nothing
            ObjCur = Nothing
        End Try
    End Function

    Public Function GetAR(ByVal vPrimaryCustomer As Double) As Decimal
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_balance_salesTableAdapter
        Dim TBL As New DseTransaksiSO.tbltrans_balance_salesDataTable
        Dim row As DseTransaksiSO.tbltrans_balance_salesRow
        Dim TipeBalance_Inv As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic
        Dim TipeBalance_SR As ClsTransaksiSalesInvoice.EnumTipeSalesBalance = ClsTransaksiSalesInvoice.EnumTipeSalesBalance.Automatic_SR
        Dim RetVal As Decimal = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_All(TBL, vPrimaryCustomer)
            For Each row In TBL.Rows
                RetVal += row.Balance
            Next

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
     ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
     Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiSOTableAdapters.tbltrans_soTableAdapter

        Try
            'tambahan find by shipto cust
            ' ahdar
            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlByNoTransaksiLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlByTanggalSampaiDari(vTanggalSampai, vTanggalDari).Value
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
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerPO_NO Then
                Return DTA.ScalarQuery_GetJmlByCustPONoLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                Return DTA.ScalarQuery_GetJmlBySalesPCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                Return DTA.ScalarQuery_GetJmlBySalesPDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                Return DTA.ScalarQuery_GetJmlByShiptoCustCodeLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                Return DTA.ScalarQuery_GetJmlByShiptoCustNameLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Sub IsiDGV_HistoricalPrice(ByVal vPrimaryCust As Double, ByVal vPrimaryProduct As Double, ByVal vDGV As DataGridView, _
    ByVal vTBL As DseTransaksiSO.DataTable6DataTable, ByVal vFrm As FrmTransaksiSO_HistorycalPrice)
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable6TableAdapter
        Dim row, row1 As DseTransaksiSO.DataTable6Row
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim DV As DataRowView
        Dim DefPrecDef As Double = -1
        Dim rowDGV As DataGridViewRow
        Dim AdaUnitMulti As Boolean = False

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGV.DataSource = Nothing
            vDGV.Rows.Clear()

            If vPrimaryCust = -1 Then
                DTA.Fill(vTBL, vPrimaryProduct)
            Else
                DTA.FillBy_PrimaryCustomer(vTBL, vPrimaryProduct, vPrimaryCust)
            End If

            vTBL.DefaultView.Sort = "[Primary Units Multi]"
            DefPrecDef = ObjProduct.GetAngkaDecimal(vPrimaryProduct)
            For ij As Double = 0 To vTBL.DefaultView.Count - 1
                row = vTBL.DefaultView(ij).Row
                row.def_prec_unit = DefPrecDef
                If row.def_prec_unit_multi = -1 Then
                    row.def_prec_unit_multi = ObjProduct.GetAngkaDecimal(vPrimaryProduct, row.Primary_Units_Multi)
                End If
                If ij <> vTBL.DefaultView.Count - 1 Then
                    row1 = vTBL.DefaultView(ij + 1).Row
                    If row.Primary_Units_Multi = row1.Primary_Units_Multi Then
                        row1.def_prec_unit_multi = row.def_prec_unit_multi
                    End If
                End If
            Next 
            For Each row In vTBL.Rows
                vFrm.dgvItem.Rows.Add()
                rowDGV = vFrm.dgvItem.Rows(vFrm.dgvItem.Rows.Count - 1)

                With rowDGV
                    .Cells(vFrm.ColAmount.Index).Value = ObjCD.DigitGrouping(row.Amount, row.Cur_Def_Prec)
                    .Cells(vFrm.ColCrrr.Index).Value = row.Currency_Code
                    .Cells(vFrm.ColDate.Index).Value = Format(row._Date, "dd MMM yyyy")
                    .Cells(vFrm.ColDisc.Index).Value = ObjCD.DigitGrouping(row.Discount, 2)
                    .Cells(vFrm.ColDiscAmount.Index).Value = ObjCD.DigitGrouping(row.Unit_Price - row.Nett_Price, row.Cur_Def_Prec)
                    .Cells(vFrm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(row.Nett_Price, row.Cur_Def_Prec)
                    .Cells(vFrm.ColNo.Index).Value = rowDGV.Index + 1
                    .Cells(vFrm.ColOrderQtyDefault.Index).Value = ObjCD.DigitGrouping(row.Qty, row.def_prec_unit)
                    .Cells(vFrm.ColOrderQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi, row.def_prec_unit_multi)
                    .Cells(vFrm.ColPONO.Index).Value = row.Customer_PO_Number
                    If row.price_by_multi_unit = 1 Then
                        .Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Units_Code_Multi
                    Else
                        .Cells(vFrm.ColPriceByMultiUnit_Str.Index).Value = "/" + row.Unit_Code
                    End If
                    .Cells(vFrm.ColPriceExclude.Index).Value = row.Price_Exclude
                    If row.Units_Code_Multi <> row.Unit_Code Then
                        rowDGV.Cells(vFrm.ColRatio.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, row.def_prec_unit_multi) & " " & row.Units_Code_Multi & _
                        "=" & ObjCD.DigitGrouping(row.Ratio_Default, row.def_prec_unit) & " " & row.Unit_Code
                        AdaUnitMulti = True
                    Else
                        rowDGV.Cells(vFrm.ColRatio.Index).Value = "-"
                    End If
                    .Cells(vFrm.ColRemarks.Index).Value = row.Remarks
                    .Cells(vFrm.ColSONo.Index).Value = row.SO_Number
                    .Cells(vFrm.ColUnitDefault.Index).Value = row.Unit_Code
                    .Cells(vFrm.ColUnitMulti.Index).Value = row.Units_Code_Multi
                    .Cells(vFrm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(row.Unit_Price, row.Cur_Def_Prec)
                End With

                'With row
                '    .Order_Qty = ObjCD.DigitGrouping(.Qty_Multi, ObjProduct.GetAngkaDecimal(vPrimaryProduct, .Primary_Units_Multi))
                '    .Unit_Code = .Units_Code_Multi
                '    .Unit_Price_Str = ObjCD.DigitGrouping(.Unit_Price, .Cur_Def_Prec)
                '    .Nett_Price_Str = ObjCD.DigitGrouping(.Nett_Price, .Cur_Def_Prec)
                '    .Amount_Str = 
                '    .Discount_Str = ObjCD.DigitGrouping(.Discount, 2)
                '    If .Customer_PO_Number = "" Then
                '        .Customer_PO_Number = .SO_Number
                '    End If
                '    ._Date = Format(CDate(._Date), "dd MMM yyyy")
                'End With
            Next

            'vDGV.DataSource = vTBL.DefaultView
            'vDGV.Columns(0).Visible = False
            'vDGV.Columns(5).Visible = False
            'vDGV.Columns(6).Visible = False
            'vDGV.Columns(8).Visible = False
            'vDGV.Columns(10).Visible = False
            'vDGV.Columns(11).Visible = False
            'vDGV.Columns(12).Visible = False
            'vDGV.Columns(13).Visible = False
            'vDGV.Columns(20).Visible = False
            'vDGV.Columns(21).Visible = False

            'vDGV.Columns(22).Visible = False
            'vDGV.Columns(23).Visible = False
            'vDGV.Columns(24).Visible = False

            'If vPrimaryCust <> -1 Then
            '    vDGV.Columns(3).Visible = False
            '    vDGV.Columns(4).Visible = False
            'End If

            'For int As Integer = 0 To vDGV.Columns.Count - 1
            '    vDGV.Columns(int).SortMode = DataGridViewColumnSortMode.NotSortable
            'Next
            'vDGV.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'vDGV.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'vDGV.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'vDGV.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'vDGV.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            'vDGV.Columns(14).HeaderText = "Unit Price"
            'vDGV.Columns(15).HeaderText = "Discount"
            'vDGV.Columns(16).HeaderText = "Nett Price"
            'vDGV.Columns(17).HeaderText = "Amount"

            'vDGV.Columns(vDGV.Columns.Count - 1).HeaderText = "Curr"
            'vDGV.Columns(vDGV.Columns.Count - 2).Visible = False
            'vDGV.Columns(vDGV.Columns.Count - 1).DisplayIndex = 14

            If AdaUnitMulti = True Then
                vFrm.ColRatio.Visible = True
                vFrm.ColOrderQtyDefault.Visible = True
                vFrm.ColUnitDefault.Visible = True
                vFrm.ColPriceByMultiUnit_Str.Visible = True
            Else
                vFrm.ColRatio.Visible = False
                vFrm.ColOrderQtyDefault.Visible = False
                vFrm.ColUnitDefault.Visible = False
                vFrm.ColPriceByMultiUnit_Str.Visible = False
            End If
            vDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjCD = Nothing
            ObjProduct = Nothing
        End Try
    End Sub

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiSO.DataTable2DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiSOTableAdapters.DataTable2TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            'tambahan find by shipto cust
            ' ahdar

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TanggalSampaiDari(TBL, vTanggalSampai, vTanggalDari)
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerPO_NO Then
                QueryResult = DTA.FillBy_CustPONOLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeCustomer Then
                QueryResult = DTA.FillBy_CustCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByNamaCustomer Then
                QueryResult = DTA.FillBy_CustNameLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonCode Then
                QueryResult = DTA.FillBy_SalesPCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySalesPersonName Then
                QueryResult = DTA.FillBy_SalesPersonDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoCode Then
                QueryResult = DTA.FillBy_ShipToCustCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByCustomerShiptoName Then
                QueryResult = DTA.FillBy_ShipToCustNameLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(9).HeaderText = "Ship to Customer Code"
                DGV.Columns(10).HeaderText = "Ship to Customer Name"
                DGV.Columns(8).DisplayIndex = 10
                DGV.Columns(9).DisplayIndex = 8
                DGV.Columns(10).DisplayIndex = 9
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
            mPrimary = -1
            mNoUrut = 0
            mNoTransaksi = ""
            mTanggalTransaksi = Now
            mPrimaryCust = -1
            mIdentifier_CustAddress = ""
            mCustContactPerson = ""
            mCustPhone = ""
            mCustPONo = ""
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
            mShipVia = ""
            mOrderBy = EnumOrderBy.PO
            mJenisFakturPajak = EnumJenisFakturPajak.FakturPajakSederhana
            mPriceCondition = EnumPriceCondition.Exclude
            mPriceConditonSama = 0
            mDiskonSama = 0
            mReferenceDiskon = 0
            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlByIdentifier(ByVal vIdentifier As String) As Decimal
        Dim ObjSO As New Inventory_Module.ClsTransaksiSalesOrder

        Try
            Return ObjSO.GetJmlProductByIdentifier2(vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjSO = Nothing
        End Try
    End Function

    Public Function GetJmlByIdentifier2(ByVal vIdentifier As String) As Decimal
        Dim ObjSO As New Inventory_Module.ClsTransaksiSalesOrder

        Try
            Return ObjSO.GetJmlProductByIdentifier3(vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjSO = Nothing
        End Try
    End Function

    Public Function GetTglByIdentifier(ByVal vIdentifier As String) As Date
        Dim DTA1 As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim vPrimaryTransaksi As Nullable(Of Double)

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            vPrimaryTransaksi = DTA1.ScalarQuery_GetPrimaryTransByIdentifier(vIdentifier)
            If vPrimaryTransaksi.HasValue = True Then
                Return Me.GetTanggal(vPrimaryTransaksi.Value)
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetNoTrans_ByIdentifier(ByVal vIdentifier As String) As String
        Dim DTA1 As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim vPrimaryTransaksi As Nullable(Of Double)

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            vPrimaryTransaksi = DTA1.ScalarQuery_GetPrimaryTransByIdentifier(vIdentifier)
            If vPrimaryTransaksi.HasValue = True Then
                Return Me.GetNoSO(vPrimaryTransaksi.Value)
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetPrimarySO_ByIdentifier(ByVal vIdentifier As String) As Double
        Dim DTA1 As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim vPrimaryTransaksi As Nullable(Of Double)

        Try
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            vPrimaryTransaksi = DTA1.ScalarQuery_GetPrimaryTransByIdentifier(vIdentifier)
            If vPrimaryTransaksi.HasValue = True Then
                Return vPrimaryTransaksi.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA1.Dispose()
        End Try
    End Function

    Public Function GetPrimaryProduct_ByIdentifier(ByVal vIdentifier As String) As Double
        Dim ObjSO As New Inventory_Module.ClsTransaksiSalesOrder

        Try
            Return ObjSO.GetPrimaryProductByIdentifier(vIdentifier)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            ObjSO = Nothing
        End Try
    End Function

    Public Sub ExplainDelivery(ByVal vIdentifier As String)
        Dim Frm As New FrmTrans_SalesDashboard_MA_DeliveryDetail
        Dim ObjSF As New Inventory_Module.ClsTransaksiShippingForecast
        Dim vTBLItemSOB As New Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryDataTable
        Dim rowDel As Inventory_Module.DseTransaksiShippingForecast.tbltrans_shipping_forecast_itemdeliveryRow
        Dim rowDGV As DataGridViewRow
        Dim DTADO As New DseTransaksiSOTableAdapters.tbltrans_do_item2TableAdapterTableAdapter
        Dim TBLDO As New DseTransaksiSO.tbltrans_do_item2TableAdapterDataTable
        Dim rowDO As DseTransaksiSO.tbltrans_do_item2TableAdapterRow
        Dim vPrimarySO, vPrimaryProduct, vPrimaryMultiUnit As Double
        Dim DTASOC As New DseTransaksiSOTableAdapters.tbltrans_soc1TableAdapter
        Dim TBLSOC As New DseTransaksiSO.tbltrans_soc1DataTable
        Dim rowSOC As DseTransaksiSO.tbltrans_soc1Row
        Dim DTASR As New DseTransaksiSOTableAdapters.tbltrans_salesreturnTableAdapter
        Dim TBLSR As New DseTransaksiSO.tbltrans_salesreturnDataTable
        Dim rowSR As DseTransaksiSO.tbltrans_salesreturnRow
        Dim DTASR1 As New DseTransaksiSOTableAdapters.tbltrans_salesreturn1TableAdapter
        Dim TBLSR1 As New DseTransaksiSO.tbltrans_salesreturn1DataTable
        Dim rowSR1 As DseTransaksiSO.tbltrans_salesreturn1Row
        Dim PrimarySRMax, PrimarySRMin As Double
        Dim MaxDate, MinDate As Date
        Dim DV As Data.DataRowView
        Dim BalanceBefore As Decimal = 0
        Dim ObjCD As New Core_Module.ClsCekData
        Dim DefPrec As Integer = 0
        Dim DefPrecMulti As Integer = 0
        Dim ObjProduct As New Core_Module.clsMasterProduct
        Dim ObjPriorityLevel As New Core_Module.ClsMasterPriorityLevel
        Dim NoSO As String = ""
        Dim ListPriorityLevel As New List(Of String)

        Try
            vPrimarySO = Me.GetPrimarySO_ByIdentifier(vIdentifier)
            vPrimaryProduct = Me.GetPrimaryProduct_ByIdentifier(vIdentifier)
            vPrimaryMultiUnit = Me.GetPrimaryUnitsMulti(vPrimarySO, vPrimaryProduct)
            DefPrec = ObjProduct.GetAngkaDecimal(vPrimaryProduct)
            DefPrecMulti = ObjProduct.GetAngkaDecimal(vPrimaryProduct, vPrimaryMultiUnit)

            DTADO.Connection.ConnectionString = Core_Module.GetConstring
            DTASOC.Connection.ConnectionString = DTADO.Connection.ConnectionString
            DTASR.Connection.ConnectionString = DTADO.Connection.ConnectionString
            DTASR1.Connection.ConnectionString = DTADO.Connection.ConnectionString

            'jelaskan shipping forecast + history delivery
            '1. isikan data history delivery
            'a. order
            Frm.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            Frm.dgvData.Rows.Clear()
            Frm.dgvData.Rows.Add()
            rowDGV = Frm.dgvData.Rows(Frm.dgvData.Rows.Count - 1)
            rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(Me.GetJmlByIdentifier2(vIdentifier), DefPrecMulti)
            rowDGV.Cells(Frm.ColDate.Index).Value = Format(Me.GetTglByIdentifier(vIdentifier), "dd MMM yyyy")
            rowDGV.Cells(Frm.ColDesc.Index).Value = "Place order"
            rowDGV.Cells(Frm.ColEdit.Index).Value = "Edit"
            rowDGV.Cells(Frm.ColNo.Index).Value = Frm.dgvData.RowCount
            rowDGV.Cells(Frm.ColPrimary.Index).Value = -1
            rowDGV.Cells(Frm.ColPrimaryPriorityLevel.Index).Value = -1
            rowDGV.Cells(Frm.ColPriorityLevel.Index).Value = ""
            rowDGV.Cells(Frm.ColQtyIn.Index).Value = "-"
            rowDGV.Cells(Frm.ColQtyOut.Index).Value = "-"
            rowDGV.Cells(Frm.ColTipeItem.Index).Value = CInt(FrmTrans_SalesDashboard_MA_DeliveryDetail.EnumTipeItem.PlaceOrder)
            rowDGV.Cells(Frm.ColDocNo.Index).Value = Me.GetNoTrans_ByIdentifier(vIdentifier)
            rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.FalseString
            rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = ""
            rowDGV.Cells(Frm.ColQtyPromiseToDeliver.Index).Value = "-"
            rowDGV.Cells(Frm.ColBookingNo.Index).Value = ""
            rowDGV.Cells(Frm.ColBookQty.Index).Value = ""
            rowDGV.DefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll

            'b. do
            DTADO.Fill(TBLDO, vPrimaryProduct, vPrimarySO)
            'c. soc
            DTASOC.Fill(TBLSOC, vPrimarySO, vPrimaryProduct)
            'd. sr
            DTASR1.Fill(TBLSR1, Inventory_Module.clsTransaksiStockInOut.EnumTipeStockInOut.StockIn_SalesReturn, vPrimaryProduct, vPrimarySO)
            MaxDate = Date.MinValue
            MinDate = Date.MaxValue

            If TBLSR1.Rows.Count > 0 Then
                PrimarySRMax = 0
                PrimarySRMin = Double.MaxValue
                For Each rowSR1 In TBLSR1.Rows
                    If PrimarySRMax < rowSR1.Primary_Transaksi Then
                        PrimarySRMax = rowSR1.Primary_Transaksi
                    End If
                    If PrimarySRMin > rowSR1.Primary_Transaksi Then
                        PrimarySRMin = rowSR1.Primary_Transaksi
                    End If
                    If MaxDate < rowSR1.Tanggal_Transaksi Then
                        MaxDate = rowSR1.Tanggal_Transaksi
                    End If
                    If MinDate > rowSR1.Tanggal_Transaksi Then
                        MinDate = rowSR1.Tanggal_Transaksi
                    End If
                Next
                DTASR.Fill(TBLSR, PrimarySRMax, PrimarySRMin)
            End If

            For Each rowDO In TBLDO.Rows
                If MaxDate < rowDO.Tanggal_Transaksi Then
                    MaxDate = rowDO.Tanggal_Transaksi
                End If
                If MinDate > rowDO.Tanggal_Transaksi Then
                    MinDate = rowDO.Tanggal_Transaksi
                End If
            Next
            For Each rowSOC In TBLSOC.Rows
                If MaxDate < rowSOC.Tanggal_Transaksi Then
                    MaxDate = rowSOC.Tanggal_Transaksi
                End If
                If MinDate > rowSOC.Tanggal_Transaksi Then
                    MinDate = rowSOC.Tanggal_Transaksi
                End If
            Next
            If TBLDO.Rows.Count > 0 Or TBLSOC.Rows.Count > 0 Or TBLSR1.Rows.Count > 0 Then
                While MinDate <= MaxDate
                    TBLDO.DefaultView.RowFilter = "[Tanggal Transaksi] = #" + Format(MinDate, "MM/dd/yyyy") + "#"
                    TBLDO.DefaultView.Sort = "[Tanggal Transaksi]"
                    For Each DV In TBLDO.DefaultView
                        rowDO = DV.Row
                        Frm.dgvData.Rows.Add()
                        rowDGV = Frm.dgvData.Rows(Frm.dgvData.Rows.Count - 1)
                        If IsNumeric(Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value) = True Then
                            BalanceBefore = Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value
                        Else
                            BalanceBefore = 0
                        End If
                        If ObjCD.DigitGrouping(BalanceBefore - rowDO.Total_Qty_Multi, DefPrecMulti) = 0 Then
                            rowDGV.Cells(Frm.ColBalance.Index).Value = "-"
                        Else
                            rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(BalanceBefore - rowDO.Total_Qty_Multi, DefPrecMulti)
                        End If

                        rowDGV.Cells(Frm.ColDate.Index).Value = Format(rowDO.Tanggal_Transaksi, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColDesc.Index).Value = "Delivery Order"
                        rowDGV.Cells(Frm.ColEdit.Index).Value = "Edit"
                        rowDGV.Cells(Frm.ColNo.Index).Value = Frm.dgvData.RowCount
                        rowDGV.Cells(Frm.ColPrimary.Index).Value = -1
                        rowDGV.Cells(Frm.ColPrimaryPriorityLevel.Index).Value = -1
                        rowDGV.Cells(Frm.ColPriorityLevel.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyIn.Index).Value = "-"
                        rowDGV.Cells(Frm.ColQtyOut.Index).Value = ObjCD.DigitGrouping(rowDO.Total_Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(Frm.ColTipeItem.Index).Value = CInt(FrmTrans_SalesDashboard_MA_DeliveryDetail.EnumTipeItem.SalesDelivery)
                        rowDGV.Cells(Frm.ColDocNo.Index).Value = rowDO.No_Transaksi
                        rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.FalseString
                        rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyPromiseToDeliver.Index).Value = "-"
                        rowDGV.Cells(Frm.ColBookingNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColBookQty.Index).Value = ""
                        rowDGV.DefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
                    Next

                    TBLSOC.DefaultView.RowFilter = "[Tanggal Transaksi] = #" + Format(MinDate, "MM/dd/yyyy") + "#"
                    TBLSOC.DefaultView.Sort = "[Tanggal Transaksi]"
                    For Each DV In TBLSOC.DefaultView
                        rowSOC = DV.Row
                        Frm.dgvData.Rows.Add()
                        rowDGV = Frm.dgvData.Rows(Frm.dgvData.Rows.Count - 1)
                        If IsNumeric(Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value) Then
                            BalanceBefore = Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value
                        Else
                            BalanceBefore = 0
                        End If

                        If ObjCD.DigitGrouping(BalanceBefore - rowSOC.Qty_Multi, DefPrecMulti) = 0 Then
                            rowDGV.Cells(Frm.ColBalance.Index).Value = "-"
                        Else
                            rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(BalanceBefore - rowSOC.Qty_Multi, DefPrecMulti)
                        End If

                        rowDGV.Cells(Frm.ColDate.Index).Value = Format(rowSOC.Tanggal_Transaksi, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColDesc.Index).Value = "Sales Order Cancellation"
                        rowDGV.Cells(Frm.ColEdit.Index).Value = "Edit"
                        rowDGV.Cells(Frm.ColNo.Index).Value = Frm.dgvData.RowCount
                        rowDGV.Cells(Frm.ColPrimary.Index).Value = -1
                        rowDGV.Cells(Frm.ColPrimaryPriorityLevel.Index).Value = -1
                        rowDGV.Cells(Frm.ColPriorityLevel.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyIn.Index).Value = "-"
                        rowDGV.Cells(Frm.ColQtyOut.Index).Value = ObjCD.DigitGrouping(rowSOC.Qty_Multi, DefPrecMulti)
                        rowDGV.Cells(Frm.ColTipeItem.Index).Value = CInt(FrmTrans_SalesDashboard_MA_DeliveryDetail.EnumTipeItem.SalesDelivery)
                        rowDGV.Cells(Frm.ColDocNo.Index).Value = rowSOC.No_Transaksi
                        rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.FalseString
                        rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyPromiseToDeliver.Index).Value = "-"
                        rowDGV.Cells(Frm.ColBookingNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColBookQty.Index).Value = ""
                        rowDGV.DefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
                    Next

                    TBLSR1.DefaultView.RowFilter = "[Tanggal Transaksi] = #" + Format(MinDate, "MM/dd/yyyy") + "#"
                    TBLSR1.DefaultView.Sort = "[Tanggal Transaksi]"
                    For Each DV In TBLSR1.DefaultView
                        rowSR1 = DV.Row
                        Frm.dgvData.Rows.Add()
                        rowDGV = Frm.dgvData.Rows(Frm.dgvData.Rows.Count - 1)
                        If IsNumeric(Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value) Then
                            BalanceBefore = Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value
                        Else
                            BalanceBefore = 0
                        End If

                        If ObjCD.DigitGrouping(BalanceBefore + rowSR1.Qty_Multi_Unit, DefPrecMulti) = 0 Then
                            rowDGV.Cells(Frm.ColBalance.Index).Value = "-"
                        Else
                            rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(BalanceBefore + rowSR1.Qty_Multi_Unit, DefPrecMulti)
                        End If

                        rowDGV.Cells(Frm.ColDate.Index).Value = Format(rowSR1.Tanggal_Transaksi, "dd MMM yyyy")
                        rowDGV.Cells(Frm.ColDesc.Index).Value = "Sales Return"
                        rowDGV.Cells(Frm.ColEdit.Index).Value = "Edit"
                        rowDGV.Cells(Frm.ColNo.Index).Value = Frm.dgvData.RowCount
                        rowDGV.Cells(Frm.ColPrimary.Index).Value = -1
                        rowDGV.Cells(Frm.ColPrimaryPriorityLevel.Index).Value = -1
                        rowDGV.Cells(Frm.ColPriorityLevel.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyIn.Index).Value = ObjCD.DigitGrouping(rowSR1.Qty_Multi_Unit, DefPrecMulti)
                        rowDGV.Cells(Frm.ColQtyOut.Index).Value = "-"
                        rowDGV.Cells(Frm.ColTipeItem.Index).Value = CInt(FrmTrans_SalesDashboard_MA_DeliveryDetail.EnumTipeItem.SalesReturn)
                        TBLSR.DefaultView.RowFilter = "[Primary Stock In Sales Return] = " + rowSR1.Primary_Transaksi.ToString
                        If TBLSR.DefaultView.Count > 0 Then
                            rowSR = TBLSR.DefaultView(0).Row
                            rowDGV.Cells(Frm.ColDocNo.Index).Value = rowSR.No_Transaksi
                        Else
                            rowDGV.Cells(Frm.ColDocNo.Index).Value = rowSR1.No_Transaksi
                        End If
                        rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.FalseString
                        rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = ""
                        rowDGV.Cells(Frm.ColQtyPromiseToDeliver.Index).Value = "-"
                        rowDGV.Cells(Frm.ColBookingNo.Index).Value = ""
                        rowDGV.Cells(Frm.ColBookQty.Index).Value = ""
                        rowDGV.DefaultCellStyle.BackColor = Core_Module.GetSMESSBackColorAll
                    Next

                    MinDate = MinDate + TimeSpan.FromDays(1)
                End While
            End If

            '2. ambil data estimate delivery
            ObjSF.HitungEstimateDelivery1(vTBLItemSOB, vIdentifier)
            NoSO = Me.GetNoSO(vPrimarySO)

            '3. isikan data estimate delivery
            For ii As Integer = vTBLItemSOB.Count - 1 To 0 Step -1
                rowDel = vTBLItemSOB.Rows(ii)
                Frm.dgvData.Rows.Add()
                rowDGV = Frm.dgvData.Rows(Frm.dgvData.Rows.Count - 1)
                If rowDGV.Index = 0 Then
                    BalanceBefore = 0
                Else
                    If IsNumeric(Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value) Then
                        BalanceBefore = Frm.dgvData.Item(Frm.ColBalance.Index, rowDGV.Index - 1).Value
                    Else
                        BalanceBefore = 0
                    End If
                End If

                rowDGV.Cells(Frm.ColBalance.Index).Value = ObjCD.DigitGrouping(BalanceBefore - rowDel.qty_available, DefPrecMulti)
                If rowDel.estimate_delivery = #1/1/1950# Then
                    rowDGV.Cells(Frm.ColDate.Index).Value = "Today"
                ElseIf rowDel.estimate_delivery = #1/1/2200# Then
                    rowDGV.Cells(Frm.ColDate.Index).Value = "n/a"
                Else
                    rowDGV.Cells(Frm.ColDate.Index).Value = Format(rowDel.estimate_delivery, "dd MMM yyyy")
                End If
                rowDGV.Cells(Frm.ColDesc.Index).Value = "Sales to " + rowDel.customer_name
                rowDGV.Cells(Frm.ColEdit.Index).Value = "Edit"
                rowDGV.Cells(Frm.ColNo.Index).Value = Frm.dgvData.Rows.Count
                rowDGV.Cells(Frm.ColPrimary.Index).Value = rowDel.primary
                rowDGV.Cells(Frm.ColPrimaryPriorityLevel.Index).Value = rowDel.primary_priority_level
                If ListPriorityLevel.Contains(ObjPriorityLevel.GetDesc(rowDel.primary_priority_level)) = False Then
                    Frm.ColPriorityLevel.Items.Add(ObjPriorityLevel.GetDesc(rowDel.primary_priority_level))
                    ListPriorityLevel.Add(ObjPriorityLevel.GetDesc(rowDel.primary_priority_level))
                End If
                rowDGV.Cells(Frm.ColPriorityLevel.Index).Value = ObjPriorityLevel.GetDesc(rowDel.primary_priority_level)
                rowDGV.Cells(Frm.ColQtyIn.Index).Value = "-"
                If rowDel.qty_available = 0 Then
                    rowDGV.Cells(Frm.ColQtyOut.Index).Value = "-"
                Else
                    rowDGV.Cells(Frm.ColQtyOut.Index).Value = ObjCD.DigitGrouping(rowDel.qty_available, DefPrecMulti)
                End If
                rowDGV.Cells(Frm.ColTipeItem.Index).Value = CInt(FrmTrans_SalesDashboard_MA_DeliveryDetail.EnumTipeItem.SOB)
                rowDGV.Cells(Frm.ColDocNo.Index).Value = NoSO

                If rowDel.estimate_delivery > rowDel.promise_to_deliver Or rowDel.qty_available < rowDel.qty_sob Then
                    rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.TrueString
                    rowDGV.DefaultCellStyle.ForeColor = Color.Red
                Else
                    rowDGV.Cells(Frm.ColIsLate.Index).Value = Boolean.FalseString
                End If
                If rowDel.promise_to_deliver = #1/1/2200# Then
                    rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = "n/a"
                Else
                    rowDGV.Cells(Frm.ColPromiseToDeliver.Index).Value = Format(rowDel.promise_to_deliver, "dd MMM yyyy")
                End If
                rowDGV.Cells(Frm.ColQtyPromiseToDeliver.Index).Value = ObjCD.DigitGrouping((ObjCD.DigitGrouping(Me.GetJmlByIdentifier2(vIdentifier), DefPrecMulti) * ObjCD.DigitGrouping(rowDel.qty_sob, DefPrec)) / ObjCD.DigitGrouping(Me.GetJmlByIdentifier(vIdentifier), DefPrec), DefPrecMulti)
                If rowDel.no_booking <> "" Then
                    rowDGV.Cells(Frm.ColBookingNo.Index).Value = rowDel.no_booking
                    rowDGV.Cells(Frm.ColBookQty.Index).Value = ObjCD.DigitGrouping(rowDel.qty_booked, DefPrec)
                Else
                    rowDGV.Cells(Frm.ColBookingNo.Index).Value = ""
                    rowDGV.Cells(Frm.ColBookQty.Index).Value = ""
                End If

                With rowDGV
                    If rowDel.qty_available > 0 Then

                        If rowDel.qty_ready_stock >= rowDel.qty_available Then
                            .Cells(Frm.ColStatus.Index).Value = "Ready"
                        ElseIf rowDel.qty_ready_stock > 0 Then
                            .Cells(Frm.ColStatus.Index).Value = "Partial"
                        Else
                            .Cells(Frm.ColStatus.Index).Value = "Waiting"
                        End If
                    Else
                        .Cells(Frm.ColStatus.Index).Value = "Waiting"
                    End If
                    If rowDel.promise_to_deliver.Date < rowDel.estimate_delivery.Date Or _
                    rowDel.qty_available < rowDel.qty_sob Then
                        rowDGV.DefaultCellStyle.ForeColor = Color.Red
                    End If
                End With
            Next

            Frm.dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            Frm.dgvData.ReadOnly = True
            Frm.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            Frm.Dispose()
            ObjSF = Nothing
            vTBLItemSOB = Nothing
            DTADO.Dispose()
            TBLDO.Dispose()
            DTASOC.Dispose()
            TBLSOC.Dispose()
            DTASR.Dispose()
            TBLSR.Dispose()
            DTASR1.Dispose()
            TBLSR1.Dispose()
            ObjCD = Nothing
            ObjProduct = Nothing
            ObjPriorityLevel = Nothing
            ListPriorityLevel = Nothing
        End Try
    End Sub

    Public Function CetakSO(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimarySO As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            Dim objCustom As New CustomSMESS.ClsTransaksiSO
            Return objCustom.CetakSO(CRV, vprimarySO, Core_Module.ModGlobalVariable.GetUser)
            objCustom = Nothing
            Exit Function
        End If

        Dim report As New CrpTransaksiSO
        Dim DS As New DseTransaksiSO
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim ObjP As New Core_Module.clsMasterProduct

        Try
            Dim DTAHeader As New DseTransaksiSOTableAdapters.TblCetakHeaderTableAdapter
            Dim RowHeader As DseTransaksiSO.TblCetakHeaderRow
            DTAHeader.FillBy_PrimarySO(DS.TblCetakHeader, vprimarySO)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSOTableAdapters.TblCetakIsiTableAdapter
            Dim RowIsi As DseTransaksiSO.TblCetakIsiRow
            DTAIsi.FillBy_PrimarySO(DS.TblCetakIsi, vprimarySO)
            For Each RowIsi In DS.TblCetakIsi
                RowIsi._Decimal = ObjP.GetAngkaDecimal(RowIsi.Primary_Product, RowIsi.Primary_Units_Multi)
            Next

            For Each RowHeader In DS.TblCetakHeader
                If RowHeader.Price_Condition = 0 Then
                    RowHeader.Price_Condition_Desc = "Exclude"
                Else
                    RowHeader.Price_Condition_Desc = "Include"
                End If

                If RowHeader.Reference_Payment_Due = 0 Then
                    RowHeader.Ref_Payment = "Sales Order Date"
                Else
                    RowHeader.Ref_Payment = "Invoice Date"
                End If
            Next

            DTAIsi.Dispose()


            report.SetDataSource(DS)
            report.SetParameterValue(report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            report.SetParameterValue(report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            report.SetParameterValue(report.Parameter_FormatYears.ParameterFieldName, FormatYears)
            report.SetParameterValue(report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
            CRV.ReportSource = report
            CRV.Show()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Public Function GetIndetifier(ByVal vPrimaryProduct As Double, ByVal vPrimarySO As Double) As String
        Dim DTASOItem As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim RetVal As String

        Try
            DTASOItem.Connection.ConnectionString = Core_Module.GetConstring
            RetVal = DTASOItem.ScalarQuery_GetIdentifier(vPrimaryProduct, vPrimarySO)
            If String.IsNullOrEmpty(RetVal) = True Then
                Return ""
            Else
                Return RetVal
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTASOItem.Dispose()
        End Try
    End Function

    Public Function GetRatioDefault(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Decimal
        Dim dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
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
        Dim dta As New DseTransaksiSOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim ratio_ As Nullable(Of Decimal) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            ratio_ = dta.ScalarQuery_GetRatioReference(vPrimaryProduct, vPrimaryUnit)
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

    Public Function GetToleranceUp(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim NoD As Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring()
            NoD = Dta.ScalarQuery_GetToleranceUp(vPrimarySO, vPrimaryProduct)
            If NoD.HasValue Then
                Return NoD.Value
            End If
            Return 0
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetPrimaryUnitsMulti(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim Value_ As New Nullable(Of Double)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Value_ = Dta.ScalarQuery_GetPrimaryUnitsMulti(vPrimarySO, vPrimaryProduct)
            If Value_.HasValue = True Then
                Return Value_.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetSORatioReference(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim Value_ As New Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Value_ = Dta.ScalarQuery_GetRatioReference(vPrimarySO, vPrimaryProduct)
            If Value_.HasValue = True Then
                Return Value_.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function GetSORatioDefault(ByVal vPrimarySO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiSOTableAdapters.tbltrans_so_itemTableAdapter
        Dim Value_ As New Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring
            Value_ = Dta.ScalarQuery_GetRatioDefault(vPrimarySO, vPrimaryProduct)
            If Value_.HasValue = True Then
                Return Value_.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            Dta.Dispose()
        End Try
    End Function

    Public Function CetakSO(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vprimarySO As Double, ByVal vUserCode As String) As Boolean
        'aul
        'karena get setting untuk cetakan so custom belum ada jadi masih di set = "" bukan = "1"
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.SO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            Dim objCustom As New CustomSMESS.ClsTransaksiSO
            objCustom.CetakSO(CRV, vprimarySO, vUserCode)
            objCustom = Nothing
            Exit Function
        End If
        ObjSetting = Nothing
        '---

        Dim report As New CrpTransaksiSO
        Dim DS As New DseTransaksiSO
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Try
            Dim DTAHeader As New DseTransaksiSOTableAdapters.TblCetakHeaderTableAdapter
            DTAHeader.FillBy_PrimarySO(DS.TblCetakHeader, vprimarySO)
            DTAHeader.Dispose()
            Dim DTAIsi As New DseTransaksiSOTableAdapters.TblCetakIsiTableAdapter
            DTAIsi.FillBy_PrimarySO(DS.TblCetakIsi, vprimarySO)
            DTAIsi.Dispose()


            report.SetDataSource(DS)
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
        End Try
    End Function

    Public Function GetHistoricalPrice(ByVal vPrimaryUnit As Double, ByVal vPrimaryProduct As Double, _
  ByVal vPrimaryCurrency As Double, ByRef vUnitPrice As Double, ByRef vDiscount As Double, _
  ByVal vPriceByMultiUnit As Integer, ByVal vPrimaryCust As Double) As Double
        Dim DTA As New DseTransaksiSO_OrderBalanceTableAdapters.tbltrans_so_item2TableAdapter
        Dim TBL As New DseTransaksiSO_OrderBalance.tbltrans_so_item2DataTable
        Dim row As DseTransaksiSO_OrderBalance.tbltrans_so_item2Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vPriceByMultiUnit = 1 Then
                If DTA.FillBy1(TBL, vPrimaryCust, vPrimaryProduct, vPrimaryUnit, vPrimaryCurrency) > 0 Then
                    row = TBL.Rows(0)
                    vUnitPrice = row.Unit_Price
                    vDiscount = row.Discount
                    'vNettPrice = row.Nett_Price
                    'vDiscAmount = row.Disc_Amount
                    'vDiscFromPerct = row.Disc_From_Perct
                Else
                    vUnitPrice = 0
                    vDiscount = 0
                    'vNettPrice = 0
                    'vDiscAmount = 0
                    'vDiscFromPerct = 1
                End If
            Else
                If DTA.Fill(TBL, vPrimaryCust, vPrimaryProduct, vPrimaryUnit, vPrimaryCurrency) > 0 Then
                    row = TBL.Rows(0)
                    vUnitPrice = row.Unit_Price
                    vDiscount = row.Discount
                    'vNettPrice = row.Nett_Price
                    'vDiscAmount = row.Disc_Amount
                    'vDiscFromPerct = row.Disc_From_Perct
                Else
                    vUnitPrice = 0
                    vDiscount = 0
                    'vNettPrice = 0
                    'vDiscAmount = 0
                    'vDiscFromPerct = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
