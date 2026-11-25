Option Explicit On
Option Strict Off

Public Class ClsTransaksiPO
    Dim ObjCD As New Core_Module.ClsCekData

    ''@rosadinnn
    ''@MUaT
    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianStockInOut_DO_Invoice
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAInOUT As New Core_Module.DseAntrianTableAdapters.tbltrans_antrianTableAdapter
    Dim ObjAntrin As New Core_Module.ClsAntrian

    Dim ObjCompanyData As New Core_Module.ClsMasterCompanyData
    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering
    Dim ObjUser As New Core_Module.ClsUser

    'dalam 1 po tidak boleh ada 2 item dengan product code sama hanya berbeda pricing saja
    'dalam 1 po, jika ada 1 product yang digunakan untuk 2 Ipo atau lebih, 
    'selain no IPO dan Qty, parameter lainnya harus sama
    Private mPrimary As Double = -1
    ''' <summary>
    ''' digunakan untuk menggenerate no transaksi secara otomatis, default value = 0
    ''' </summary>
    ''' <remarks></remarks>
    Private mPrimarySupp As Double = -1
    Private mNoTransaksi As String = ""
    Private mNoUrut As Double = 0
    Private mTanggalTransaksi As Date = Now
    Private mIdentifier_SuppAddress As String = ""
    Private mSuppContactPerson As String = ""
    Private mSuppPhoneNo As String = ""
    Private mPaymentTerm As String = ""
    Private mShipVia As String = ""
    Private mPaymentDue As Integer = 0
    Private mReferencePaymentDue As Integer = 0
    Private mPrimaryPurchaseStaffAuthorization As Double = -1
    Private mPurchaseStaffAuthorizationDate As Date = Now
    Private mPrimaryManagerStaffAuthorization As Double = -1
    Private mPurchaseManagerAuthorizationDate As Date = Now
    Private mPrimaryDirectorStaffAuthorization As Double = -1
    Private mPurchaseDirectorAuthorizationDate As Date = Now
    Private mAuthorizationlevel As Integer = 0
    Private mAuthorizationNeeded As Integer = 0
    Private mPrimaryCurrency As Double = 1
    Private mSubTotal As Double = 0
    Private mTaxExclude As Boolean = False
    Private mTax As Double = 0
    Private mTaxAmount As Double = 0
    Private mTotal As Double = 0
    Private mRemarks As String = ""
    Private mPreparedBy As String = ""
    Private mPreparedBy_Jabatan As String = ""
    Private mApprovedBy As String = ""
    Private mApprovedBy_Jabatan As String = ""
    Private mAcknowledgeBy As String = ""
    Private mAcknowlegdeBy_Jabatan As String = ""
    Private mIdentifierShipTo As String = Me.ObjCompanyData.GetIdentifierShipToDefault
    Private mNeedDP As Integer = 0
    Private mDPPercent As Decimal = 0
    Private mDPAmount As Decimal = 0
    Private mDPFromPercent As Integer = 1

    Dim mPrintedBy As String = ""
    Dim mPrintedDate As Date = Now
    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    Public Enum EnumPriceCondition As Integer
        Exclude
        Include
    End Enum

    Public Enum EnumRefPaymentDate As Integer
        InvoiceDate
        PODate
        BillOfLeadingDate
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

    Public Enum Refenrence_Payment_Due As Integer
        PO_Date
        Invoice_Date
    End Enum

    Public Enum EnumStatusPO As Integer
        AllPO
        AuthorizeLvl_1
        AuthorizeLvl_2
        AuthorizeLvl_3
        UnauthorizedOnly
    End Enum

    Public Enum EnumStatusPrintPO As Integer
        AllPO
        PrintedPOOnly
        UnPrintedPOOnly
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

    Public Property PreparedBy() As String
        Get
            Return Me.mPreparedBy
        End Get
        Set(ByVal value As String)
            mPreparedBy = value
        End Set
    End Property

    Public Property IdentifierShipTo() As String
        Get
            Return Me.mIdentifierShipTo
        End Get
        Set(ByVal value As String)
            mIdentifierShipTo = value
        End Set
    End Property

    Public Property PreparedBy_Jabatan() As String
        Get
            Return Me.mPreparedBy_Jabatan
        End Get
        Set(ByVal value As String)
            mPreparedBy_Jabatan = value
        End Set
    End Property

    Public Property ApprovedBy() As String
        Get
            Return Me.mApprovedBy
        End Get
        Set(ByVal value As String)
            mApprovedBy = value
        End Set
    End Property

    Public Property ApprovedBy_Jabatan() As String
        Get
            Return Me.mApprovedBy_Jabatan
        End Get
        Set(ByVal value As String)
            mApprovedBy_Jabatan = value
        End Set
    End Property

    Public Property AcknowledgeBy() As String
        Get
            Return Me.mAcknowledgeBy
        End Get
        Set(ByVal value As String)
            mAcknowledgeBy = value
        End Set
    End Property

    Public Property AcknowledgeBy_Jabatan() As String
        Get
            Return Me.mAcknowlegdeBy_Jabatan
        End Get
        Set(ByVal value As String)
            mAcknowlegdeBy_Jabatan = value
        End Set
    End Property

    Public Property primary() As Double
        Get
            Return Me.mPrimary
        End Get
        Set(ByVal value As Double)
            mPrimary = value
        End Set
    End Property

    Public Property primarySupplier() As Double
        Get
            Return Me.mPrimarySupp
        End Get
        Set(ByVal value As Double)
            mPrimarySupp = value
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

    Public Property urutan() As Double
        Get
            Return Me.mNoUrut
        End Get
        Set(ByVal value As Double)
            mNoUrut = value
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

    Public Property IdentifierSuppAddress() As String
        Get
            Return Me.mIdentifier_SuppAddress
        End Get
        Set(ByVal value As String)
            mIdentifier_SuppAddress = value
        End Set
    End Property

    Public Property SuppContactPerson() As String
        Get
            Return Me.mSuppContactPerson
        End Get
        Set(ByVal value As String)
            mSuppContactPerson = value
        End Set
    End Property
    Public Property SuppPhoneNo() As String
        Get
            Return Me.mSuppPhoneNo
        End Get
        Set(ByVal value As String)
            mSuppPhoneNo = value
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

    Public Property ShipVia() As String
        Get
            Return Me.mShipVia
        End Get
        Set(ByVal value As String)
            mShipVia = value
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

    Public Property ReferencePaymentDue() As Integer
        Get
            Return Me.mReferencePaymentDue
        End Get
        Set(ByVal value As Integer)
            mReferencePaymentDue = value
        End Set
    End Property

    Public Property PrimaryPurchaseStaffAuthorization() As Double
        Get
            Return Me.mPrimaryPurchaseStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryPurchaseStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseStaffAuthorizationDate() As Date
        Get
            Return Me.mPurchaseStaffAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseStaffAuthorizationDate = value
        End Set
    End Property

    Public Property PrimaryManagerStaffAuthorization() As Double
        Get
            Return Me.mPrimaryManagerStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryManagerStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseManagerAuthorizationDate() As Date
        Get
            Return Me.mPurchaseManagerAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseManagerAuthorizationDate = value
        End Set
    End Property

    Public Property PrimaryDirectorStaffAuthorization() As Double
        Get
            Return Me.mPrimaryDirectorStaffAuthorization
        End Get
        Set(ByVal value As Double)
            mPrimaryDirectorStaffAuthorization = value
        End Set
    End Property

    Public Property PurchaseDirectorAuthorizationDate() As Date
        Get
            Return Me.mPurchaseDirectorAuthorizationDate
        End Get
        Set(ByVal value As Date)
            mPurchaseDirectorAuthorizationDate = value
        End Set
    End Property

    Public Property Authorizationlevel() As Integer
        Get
            Return Me.mAuthorizationlevel
        End Get
        Set(ByVal value As Integer)
            mAuthorizationlevel = value
        End Set
    End Property

    Public Property AuthorizationNeeded() As Integer
        Get
            Return Me.mAuthorizationNeeded
        End Get
        Set(ByVal value As Integer)
            mAuthorizationNeeded = value
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

    Public Property SubTotal() As Double
        Get
            Return Me.mSubTotal
        End Get
        Set(ByVal value As Double)
            mSubTotal = value
        End Set
    End Property

    Public Property TaxExclude() As Boolean
        Get
            Return Me.mTaxExclude
        End Get
        Set(ByVal value As Boolean)
            mTaxExclude = value
        End Set
    End Property

    Public Property Tax() As Double
        Get
            Return Me.mTax
        End Get
        Set(ByVal value As Double)
            mTax = value
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

    Public Property Total() As Double
        Get
            Return Me.mTotal
        End Get
        Set(ByVal value As Double)
            mTotal = value
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

    Public Property CreatedBy() As String
        Get
            Return mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property PrintedBy() As String
        Get
            Return mPrintedBy
        End Get
        Set(ByVal value As String)
            mPrintedBy = value
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

    Public Property PrintedDate() As Date
        Get
            Return mPrintedDate
        End Get
        Set(ByVal value As Date)
            mPrintedDate = value
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

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByNoTrans(vNoTransaksi).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByNoTrans(vNoTransaksi).Value > 0 Then
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

    Public Overridable Function ApakahSudahAda(ByVal vNoTransaksi As String, ByVal vPrimaryAOPQuick As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJml2(vNoTransaksi, vPrimaryAOPQuick).HasValue = True Then
                If DTA.ScalarQuery_GetJml2(vNoTransaksi, vPrimaryAOPQuick).Value > 0 Then
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

    Public Function GetHistoricalPrice(ByVal vPrimaryUnit As Double, ByVal vPrimaryProduct As Double, _
    ByVal vPrimaryCurrency As Double, ByRef vUnitPrice As Double, ByRef vDiscount As Double, _
    ByRef vNettPrice As Double, ByRef vDiscAmount As Decimal, ByRef vDiscFromPerct As Integer, ByVal vPriceByMultiUnit As Integer) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.DataTable8TableAdapter
        Dim TBL As New DseTransaksiPO.DataTable8DataTable
        Dim row As DseTransaksiPO.DataTable8Row

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If vPriceByMultiUnit = 1 Then
                If DTA.Fill(TBL, vPrimaryProduct, vPrimaryCurrency, vPrimaryUnit) > 0 Then
                    row = TBL.Rows(0)
                    vUnitPrice = row.Unit_Price
                    vDiscount = row.Discount
                    vNettPrice = row.Nett_Price
                    vDiscAmount = row.Disc_Amount
                    vDiscFromPerct = row.Disc_From_Perct
                Else
                    vUnitPrice = 0
                    vDiscount = 0
                    vNettPrice = 0
                    vDiscAmount = 0
                    vDiscFromPerct = 1
                End If
            Else
                If DTA.FillBy1(TBL, vPrimaryProduct, vPrimaryCurrency, vPrimaryUnit) > 0 Then
                    row = TBL.Rows(0)
                    vUnitPrice = row.Unit_Price
                    vDiscount = row.Discount
                    vNettPrice = row.Nett_Price
                    vDiscAmount = row.Disc_Amount
                    vDiscFromPerct = row.Disc_From_Perct
                Else
                    vUnitPrice = 0
                    vDiscount = 0
                    vNettPrice = 0
                    vDiscAmount = 0
                    vDiscFromPerct = 1
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function ExplainHistoricalPrice(ByVal vPrimaryProduct As Double, ByVal _
    TBL As DseTransaksiPO.DataTable9DataTable, ByVal vDGV As DataGridView, _
    ByVal ObjCD As Core_Module.ClsCekData) As Double
        ''@rosadinnn
        ''@MUaT
        Dim DTA As New DseTransaksiPOTableAdapters.DataTable9TableAdapter
        Dim row As DseTransaksiPO.DataTable9Row
        Dim rowDgv As DataGridViewRow
        Dim frm As New FrmTransaksiPO_HistoricalPrice
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objCurr As New Core_Module.clsMasterCurrency
        Dim DefPrec As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            frm.ColTanggal.HeaderText = "Date"

            vDGV.DataSource = Nothing
            If DTA.FillBy_PrimaryProduct(TBL, vPrimaryProduct) > 0 Then
                For Each row In TBL.Rows
                    row.Unit_Price_Str = ObjCD.DigitGrouping(row.Unit_Price)
                    row.Discount_Str = ObjCD.DigitGrouping(row.Discount)
                    row.Nett_Price_Str = ObjCD.DigitGrouping(row.Nett_Price)
                    row.PO_Qty_Str = ObjCD.DigitGrouping(row.PO_Qty)

                    vDGV.Rows.Add()
                    rowDgv = vDGV.Rows(vDGV.Rows.Count - 1)
                    With rowDgv
                        .Cells(frm.ColNo.Index).Value = vDGV.Rows.Count.ToString
                        .Cells(frm.ColTanggal.Index).Value = Format(row.Tanggal_Transaksi, "dd MMM yyyy")
                        .Cells(frm.ColPONumber.Index).Value = row.PO_Number
                        .Cells(frm.ColSupplierCode.Index).Value = row.Supplier_Code
                        .Cells(frm.ColSupplierName.Index).Value = row.Supplier_Name
                        .Cells(frm.ColCurrencyCode.Index).Value = row.Currency_Code
                        DefPrec = objCurr.GetDefaultPrecission(row.Primary_Currency)
                        .Cells(frm.ColUnitPrice.Index).Value = ObjCD.DigitGrouping(row.Unit_Price, DefPrec)
                        .Cells(frm.ColDiscount.Index).Value = ObjCD.DigitGrouping(row.Discount, 2)
                        .Cells(frm.ColNettPrice.Index).Value = ObjCD.DigitGrouping(row.Nett_Price, DefPrec)
                        .Cells(frm.ColQtyMulti.Index).Value = ObjCD.DigitGrouping(row.Qty_Multi, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi))
                        .Cells(frm.ColUnitCodeMulti.Index).Value = row.Unit_Code_Multi
                        .Cells(frm.ColRatio.Index).Value = ObjCD.DigitGrouping(row.Ratio_Reference, objProduct.GetAngkaDecimal(row.Primary_Product, row.Primary_Unit_Multi)) & " " & row.Unit_Code_Multi & " = " & ObjCD.DigitGrouping(row.Ratio_Default, objProduct.GetAngkaDecimal(row.Primary_Product)) & " " & row.Unit_Code
                        .Cells(frm.ColQty.Index).Value = ObjCD.DigitGrouping(row.PO_Qty, objProduct.GetAngkaDecimal(row.Primary_Product))
                        .Cells(frm.ColUnitCode.Index).Value = row.Unit_Code
                        If row.price_by_multi_unit = 0 Then
                            .Cells(frm.ColPriceByMultiUnit_str.Index).Value = "/" + .Cells(frm.ColUnitCode.Index).Value
                        Else
                            .Cells(frm.ColPriceByMultiUnit_str.Index).Value = "/" + .Cells(frm.ColUnitCodeMulti.Index).Value
                        End If
                    End With
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            objProduct = Nothing
            objCurr = Nothing
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada PO dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl1(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl1(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl1(vPrimaryPO).Value
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
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada PO dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl2(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl2(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl2(vPrimaryPO).Value
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
    ''' mengembalikan nilai primary user yang melakukan proses otorisasi pada PO dan level yang bersangkutan
    ''' </summary>
    ''' <param name="vPrimaryPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetAuthorizeBy_Lvl3(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl3(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryUser_AuthorizeLvl3(vPrimaryPO).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAuthorizationLevel(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Return DTA.ScalarQuery_getAuthorizationLevel(vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAuthorizationLevelNeeded(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetAuthorizationLevelNeeded(vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahOtorisasi(ByVal vPrimaryTransaksi As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlAuthorizedPO(vPrimaryTransaksi).Value > 0 Then
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

    Private Function ApakahSudahAda(ByVal vPrimaryPO As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).HasValue = True Then
                If DTA.ScalarQuery_GetJmlByPrimaryPO(vPrimaryPO).Value = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
       ByVal vSalesStaff As String, ByVal StrictEditting As Boolean, _
       Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim CreatedBy As String = ""
        Dim ObjUser As New Core_Module.ClsUser

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'jika sudah melewati batas penguncian transaksi maka tidak bisa edit data
            'msgbox("belum")
            'batasan edit PO:
            '1. Jika sudah di otorisasi, maka tidak bisa diedit (KECUALI USER TERTENTU)
            If Me.IsSkipCheckAuthrization1(vSalesStaff, ObjUser) = False Then
                If DTA.ScalarQuery_getAuthorizationLevel(vPrimaryTransaksi).HasValue = False Then
                Else
                    If DTA.ScalarQuery_getAuthorizationLevel(vPrimaryTransaksi).Value <> 0 Then
                        If TampilkanPesan = True Then
                            MsgBox("This transaction have been authorize, please cancel authorization first.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If


            '2.a. po ini sudah tidak ada didalam database, minta untuk direload
            If Me.ApakahSudahAda(vPrimaryTransaksi) = False Then
                MsgBox("Other user have modified or delete this data, please reload data.", _
                MsgBoxStyle.Information, "Information")
            End If

            '2. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            If StrictEditting = True Then
                CreatedBy = DTA.ScalarQuery_getCreatedBy(vPrimaryTransaksi)
                If CreatedBy <> vSalesStaff Then
                    If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiPO_Override_Edit") = False Then
                        If TampilkanPesan = True Then
                            MsgBox("Only user with ID:" + CreatedBy + " can edit this data.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    End If
                End If
            End If

            'po ini adalah po yang diciptakan untuk quick purchase
            If DTA.ScalarQuery_GetPrimaryAOPQuick(vPrimaryTransaksi).HasValue = True Then
                Dim ObjAOP As New ClsTransaksiAOP
                Dim PrimaryAOP As Double = DTA.ScalarQuery_GetPrimaryAOPQuick(vPrimaryTransaksi).Value
                If PrimaryAOP <> -1 Then
                    MsgBox("This PO is refered to quick purchase: " + _
                    ObjAOP.GetNoTransaksi(PrimaryAOP) + ". You cannot edit this data.", MsgBoxStyle.Information, "Information")
                    ObjAOP = Nothing
                    Return False
                End If
                ObjAOP = Nothing
            End If

            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjUser = Nothing
        End Try
    End Function

    Public Function IsSkipCheckAuthrization1(ByVal vUserCode As String, ByVal ObjUser As Core_Module.ClsUser, Optional ByVal vShowWarning As Boolean = True) As Boolean
        Try
            If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPO_Override_Otorisasi") = False Then
                Return False
            Else
                Return True
            End If

            Return False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean, Optional ByVal vCekDataBenar As Boolean = True) As Boolean
        Try
            If vCekDataBenar = False Then
                Return True
            End If

            'msgbox("belum")
            'data salah bila:
            '0. untuk penomoran, yakinkan sistem penomoran po lokal dan import sama2 autonumber atau sama2 manual
            If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) <> _
            Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor) Then
                MsgBox("Local dan import po, must have identical numbering system.", MsgBoxStyle.Information, "Information")
                Exit Function
            End If

            '1. Sudah ada no transaksi yang sama sebelumnya (untuk transaksi baru dan edit BEDA)
            '       untuk transaksi baru - dan autonumber-langsung ambil nomor baru
            '       untuk transaksi edit - dan autonumber-dianggap sama karena memang no transaksi tidak bisa diedit.
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                        Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupp)
                    Else
                        MsgBox("PO Number duplication, please change PO number.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            Else
                If Me.ApakahSudahAda(Me.mNoTransaksi) = True Then
                    If Me.mPrimary <> Me.GetPrimary(Me.mNoTransaksi) Then
                        If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                            Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupp)
                        Else
                            MsgBox("PO Number duplication, please change PO number.", MsgBoxStyle.Information, "Information")
                            Return False
                        End If
                    End If
                End If
            End If

            '4. jika tidak ada item yang akan disimpan

            '5. jika item yang disimpan mengakibatkan order balance minus

            '6. jika item yang disimpan nilainya lebih kecil dari yang sudah dibooking via IPO--> hati - hati ini

            '7. jika item yang disimpan nilainya lebih kecil dari yang sudah booking inv
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetNoBaru1(ByVal vTanggalTransaksi As Date, Optional ByVal vPrimarySupplier As Double = -1) As String
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut As Double = 0
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ApakahLokal As Boolean = ObjSupplier.ApakahLokal(vPrimarySupplier)
        Dim ForeignSupplier As Boolean = True
        Dim ObjCD As New Core_Module.ClsCekData
        Dim RetVal As String = ""
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari

            If ApakahLokal = True Then
                ForeignSupplier = False
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal
            Else
                ForeignSupplier = True
                TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor
            End If
            Me.ObjNumbering.GetTanggalDariSampai(TanggalStart, TanggalEnd, TipeTransaksi, vTanggalTransaksi)

            'cari no urut
            If DTA.ScalarQuery_MAXUrutan(TanggalEnd, TanggalStart, ForeignSupplier).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_MAXUrutan(TanggalEnd, TanggalStart, ForeignSupplier).Value + 1)
                If NoUrut <= 0 Then
                    NoUrut = 1
                End If
            Else
                NoUrut = 1
            End If

            'cari nomor baru
            RetVal = Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)
            While Me.ApakahSudahAda(RetVal) = True
                NoUrut = NoUrut + 1
                RetVal = Me.ObjNumbering.GetNomorTransaksiBaru1(vTanggalTransaksi, NoUrut, TipeTransaksi)
            End While

            Return RetVal
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Function GetNoBaru_Obs(ByVal vTanggalTransaksi As Date, Optional ByVal vPrimarySupplier As Double = -1) As String
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim TanggalStart, TanggalEnd As Date
        Dim NoUrut, Bulan As String
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim ApakahLokal As Boolean = ObjSupplier.ApakahLokal(vPrimarySupplier)
        Dim TipeStr As String = ""
        Dim ForeignSupplier As Boolean = True
        Dim ObjCD As New Core_Module.ClsCekData

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'cari tanggal sampai dan dari
            TanggalEnd = ObjCD.GetLastDateOfTheYear(vTanggalTransaksi)
            TanggalStart = ObjCD.GetFisrtDateOfTheYear(vTanggalTransaksi)

            If ApakahLokal = True Then
                TipeStr = "L"
                ForeignSupplier = False
            Else
                TipeStr = "I"
                ForeignSupplier = True
            End If

            'cari no urut
            If DTA.ScalarQuery_MAXUrutan(TanggalEnd, TanggalStart, ForeignSupplier).HasValue = True Then
                NoUrut = CStr(DTA.ScalarQuery_MAXUrutan(TanggalEnd, TanggalStart, ForeignSupplier).Value + 1)
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

            While Me.ApakahSudahAda(NoUrut + "/" + "PO" + "/" + TipeStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)) = True
                NoUrut = CStr(CDbl(NoUrut) + 1)
                While NoUrut.Length < 6
                    NoUrut = "0" + NoUrut
                End While
            End While

            Return NoUrut + "/" + "PO" + "/" + TipeStr + "/" + Bulan + "/" + CStr(vTanggalTransaksi.Year)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            ObjCD = Nothing
        End Try
    End Function

    Public Overridable Function CariDetail(ByVal vPrimaryTransaksi As Double, _
      ByVal vDgvItem As DataGridView, ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable, ByVal vFrm As FrmTransaksiPO) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_NewTableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_po_NewDataTable
        Dim row As DseTransaksiPO.tbltrans_po_NewRow
        'Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        'Dim TBL As New DseTransaksiPO.tbltrans_poDataTable
        'Dim row As DseTransaksiPO.tbltrans_poRow
        Dim ObjItem As New ClsTransaksiPOItem
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'fil basis data
            If DTA.Fill(TBL, vPrimaryTransaksi) > 0 Then
                'salin data header
                row = TBL.Rows(0)

                With row
                    Me.mPrimary = vPrimaryTransaksi
                    Me.mPrimarySupp = .Primary_Supplier
                    Me.mNoTransaksi = .No_Transaksi
                    Me.mNoUrut = .Ururtan
                    Me.mTanggalTransaksi = .Tanggal_Transaksi
                    Me.mIdentifier_SuppAddress = .Indentifier_Alamat_Supplier
                    Me.mSuppContactPerson = .Contact_Person
                    Me.mSuppPhoneNo = .Supp_Phone_No
                    Me.mPaymentTerm = .Payment_Term
                    Me.mShipVia = .Ship_Via
                    Me.mPaymentDue = .Payment_Due
                    Me.mReferencePaymentDue = .Reference_Payment_Due
                    Me.mPrimaryPurchaseStaffAuthorization = .Primary_Purchase_Staff_Authorization
                    Me.mPurchaseStaffAuthorizationDate = .Purchase_Staff_Authorization_Date
                    Me.mPrimaryManagerStaffAuthorization = .Primary_Purchase_Manager_Authorization
                    Me.mPurchaseManagerAuthorizationDate = .Purchase_Manager_Authorization_Date
                    Me.mPrimaryDirectorStaffAuthorization = .Primary_Director_Authorization
                    Me.mPurchaseDirectorAuthorizationDate = .Director_Authorization_Date
                    Me.mAuthorizationlevel = .Authorization_Level
                    Me.mAuthorizationNeeded = .Authorization_Level_Needed
                    Me.mPrimaryCurrency = .Primary_Currency
                    Me.mSubTotal = .Sub_Total
                    Me.mTaxExclude = .Tax_Exclude
                    Me.mTax = .Tax
                    Me.mTaxAmount = .Tax_Amount
                    Me.mTotal = .Total
                    Me.mRemarks = .Remarks
                    Me.mApprovedBy = .ApprovedByName
                    Me.mApprovedBy_Jabatan = .ApprovedByPosition
                    Me.mPreparedBy = .PreparedByName
                    Me.mPreparedBy_Jabatan = .PreparedByPosition
                    Me.mAcknowledgeBy = .AcknowledgeByName
                    Me.mAcknowlegdeBy_Jabatan = .AcknowledgeByPostion
                    Me.mIdentifierShipTo = .Identifier_Ship_To
                    Me.mNeedDP = .need_dp
                    Me.mDPAmount = .dp_amount
                    Me.mDPFromPercent = .dp_from_percent
                    Me.mDPPercent = .dp_percent

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mPrintedBy = .Printed_by
                    Me.mPrintedDate = .Printed_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With
                'salin data item 
                ObjItem.IsiDGV(vDgvItem, vPrimaryTransaksi, ObjCur.GetDefaultPrecission(Me.mPrimaryCurrency), vTBLItemDelivery, vFrm, Me)
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

    Public Sub ResetObject(ByVal vDgvItem As DataGridView)
        Try
            mPrimarySupp = -1
            mNoTransaksi = ""
            mNoUrut = 0
            mTanggalTransaksi = Now
            mIdentifier_SuppAddress = ""
            mSuppContactPerson = ""
            mSuppPhoneNo = ""
            mPaymentTerm = ""
            mShipVia = ""
            mPaymentDue = 0
            mReferencePaymentDue = 0
            mPrimaryPurchaseStaffAuthorization = -1
            mPurchaseStaffAuthorizationDate = Now
            mPrimaryManagerStaffAuthorization = -1
            mPurchaseManagerAuthorizationDate = Now
            mPrimaryDirectorStaffAuthorization = -1
            mPurchaseDirectorAuthorizationDate = Now
            mAuthorizationlevel = 0
            mAuthorizationNeeded = 3
            mPrimaryCurrency = 1
            mSubTotal = 0
            mTaxExclude = 0
            mTax = 0
            mTaxAmount = 0
            mTotal = 0
            mRemarks = ""
            Me.mPreparedBy = ""
            Me.mPreparedBy_Jabatan = ""
            Me.mApprovedBy_Jabatan = ""
            Me.mApprovedBy = ""
            Me.mAcknowledgeBy = ""
            Me.mAcknowlegdeBy_Jabatan = ""

            vDgvItem.Rows.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Overridable Function SimpanBaru(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vfrm As FrmTransaksiPO, _
            ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable, Optional ByVal vSETPOB As Boolean = True) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Dim SW As New Stopwatch

        Try
            SW.Start()
            'ambil no antrian
            Me.ObjAntrin.GetNoAntrian(Me.vIdentifier, Me.ObjRnd, Me.DTAInOUT, Me.vTipeAntrian, Me.vPrimary_Antrian, Me.primary, _
            Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseOrder_SimpanBaru, vUserCode)
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
                            retval = Me.SimpanBaru1(vdgvItem, vUserCode, vfrm, vTBLItemDelivery)
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

    Public Overridable Function SimpanBaru1(ByVal vdgvItem As DataGridView, ByVal vUserCode As String, ByVal vfrm As FrmTransaksiPO, _
    ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable, Optional ByVal vSETPOB As Boolean = True, _
    Optional ByVal vCekDataBenar As Boolean = True, Optional ByVal vFromCutOff As Boolean = False) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim ObjItem As New ClsTransaksiPOItem
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal
        Dim ObjSupplier As New Core_Module.clsMasterSupplier

        Dim CreatedBy_ As String = vUserCode
        Dim CreatedDate_ As Date = Now
        Dim ModifiedBy_ As String = vUserCode
        Dim ModifiedDate_ As Date = Now
        Dim TimeStamp_ As Date = Now

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(True, vCekDataBenar) = True Then
                'hitung level otorisasi yang dibutuhkan
                Me.mAuthorizationNeeded = 3
                'insert query
                Dim TaxExclude As Integer = 0
                If Me.TaxExclude = True Then
                    TaxExclude = 1
                End If

                If ObjSupplier.ApakahLokal(Me.mPrimarySupp) = True Then
                    TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal
                Else
                    TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor
                End If
                Me.mNoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)

                If vFromCutOff = True Then
                    CreatedBy_ = Me.mCreatedBy
                    CreatedDate_ = Me.mCreatedDate
                    ModifiedBy_ = Me.mModifiedBy
                    ModifiedDate_ = Me.mModifiedDate
                    TimeStamp_ = Me.mTimestamp
                End If

                If DTA.Insert(Me.primarySupplier, Me.NoTransaksi, Me.urutan, Me.TanggalTransaksi, _
                Me.IdentifierSuppAddress, Me.SuppContactPerson, Me.SuppPhoneNo, Me.PaymentTerm, _
                Me.ShipVia, Me.PaymentDue, Me.ReferencePaymentDue, Me.PrimaryPurchaseStaffAuthorization, _
                Me.PurchaseStaffAuthorizationDate, Me.PrimaryManagerStaffAuthorization, Me.PurchaseManagerAuthorizationDate, _
                Me.PrimaryDirectorStaffAuthorization, Me.PurchaseDirectorAuthorizationDate, Me.Authorizationlevel, Me.AuthorizationNeeded, _
                Me.PrimaryCurrency, Me.SubTotal, TaxExclude, Me.Tax, Me.TaxAmount, Me.Total, Me.Remarks, CreatedBy_, CreatedDate_, _
                ModifiedBy_, ModifiedDate_, TimeStamp_, Me.mPreparedBy, Me.mPreparedBy_Jabatan, Me.mApprovedBy, Me.mApprovedBy_Jabatan, _
                Me.mAcknowledgeBy, Me.mAcknowlegdeBy_Jabatan, Me.mIdentifierShipTo, Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent) > 0 Then
                    'update data yang berubah setelah insert query
                    Me.mPrimary = DTA.ScalarQuery_GetPrimaryByNotransaksi(Me.mNoTransaksi).Value
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

                    'simpan data item baru
                    ObjItem.Simpan(vdgvItem, Me.mPrimary, vfrm, vTBLItemDelivery)
                    'set purchase order balance
                    If vSETPOB = True Then
                        ObjItem.SetOrderBalance(Me.mPrimary)
                    End If
                    'cek kembali nilai penomoran untuk bulan yang berbeda
                    'If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                    '    Me.CekNomor1(DTA)
                    'End If

                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Or _
            ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 'No Transaksi'" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupp)
                    Return Me.SimpanBaru1(vdgvItem, vUserCode, vfrm, vTBLItemDelivery, vSETPOB)
                Else
                    MsgBox("PO Number duplication, please change PO number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
            MsgBox("Error on po")
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjSupplier = Nothing
        End Try
    End Function

    Public Sub SetPrimaryQuickAOP(ByVal vPrimaryAOP As Double, ByVal vPrimaryPO As Double)
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.UpdateQuery_SetQuickAOP(vPrimaryAOP, vPrimaryPO)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    'Private Sub CekNomor1(ByVal DTA As DseTransaksiPOTableAdapters.tbltrans_poTableAdapter)
    '    Dim TanggalDari, TanggalSampai As Date
    '    Dim Urutan As Integer = Me.urutan
    '    Dim LoopGo As Boolean = True
    '    Dim JmlTrans As Integer
    '    Dim ObjSupp As New Core_Module.clsMasterSupplier
    '    Dim Foreign, LoopCount As Integer
    '    Dim ObjCD As New Core_Module.ClsCekData
    '    Dim PrimaryMax As Double = -1
    '    Dim ObjNumbering As New Core_Module.ClsMasterTransactionNumbering

    '    Try
    '        DTA.Connection.ConnectionString = Core_Module.GetConstring

    '        If ObjSupp.ApakahLokal(Me.mPrimarySupp) = True Then
    '            Foreign = 0
    '        Else
    '            Foreign = 1
    '        End If
    '        TanggalDari = ObjCD.GetFisrtDateOfTheYear(Me.mTanggalTransaksi)
    '        If TanggalDari.Date < #5/1/2010# Then
    '            TanggalDari = #5/1/2010#
    '        End If
    '        TanggalSampai = ObjCD.GetLastDateOfTheYear(Me.mTanggalTransaksi)

    '        LoopCount = 0
    '        While LoopGo = True

    '            If DTA.ScalarQuery_GetJmlByUrutan(Me.urutan, TanggalSampai, TanggalDari, Foreign).HasValue = True Then
    '                JmlTrans = DTA.ScalarQuery_GetJmlByUrutan(Me.urutan, TanggalSampai, TanggalDari, Foreign).Value
    '            Else
    '                JmlTrans = 1
    '            End If
    '            If JmlTrans <= 1 Then
    '                LoopGo = False
    '            Else
    '                If DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, Foreign).HasValue = True Then
    '                    PrimaryMax = DTA.ScalarQuery_GetPrimaryMaxByUrutan(Me.mNoUrut, TanggalSampai, TanggalDari, Foreign)
    '                Else
    '                    PrimaryMax = -1
    '                End If
    '                If PrimaryMax = Me.mPrimary Then
    '                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupp)
    '                    If Foreign = 0 Then
    '                        Me.mNoUrut = ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal, Me.mTanggalTransaksi)
    '                    Else
    '                        Me.mNoUrut = ObjNumbering.GetUrutan(Me.mNoTransaksi, Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor, Me.mTanggalTransaksi)
    '                    End If
    '                End If
    '            End If
    '            LoopCount += 1
    '        End While

    '        If LoopCount > 1 Then
    '            DTA.UpdateQuery_NoTrans_Urutan(Me.mNoUrut, Me.mNoTransaksi, Me.mPrimary)
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    Finally
    '        ObjSupp = Nothing
    '        ObjCD = Nothing
    '    End Try
    'End Sub

    Public Function GetTanggalPO(ByVal vPrimaryPO As Double) As Date
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTanggal(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetTanggal(vPrimaryPO).Value
            Else
                Return #1/1/1950#
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPO, ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable) As Boolean
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
            Me.vPrimary_Antrian, Me.mPrimary, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseOrder_SimpanPerubahan, vUserCode)


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
                            retval = Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm, vTBLItemDelivery)
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

    Public Overridable Function SimpanPerubahan1(ByVal dgvItem As DataGridView, ByVal vUserCode As String, ByVal vFrm As FrmTransaksiPO, ByVal vTBLItemDelivery As DseTransaksiPO.tbltrans_po_itemdeliveryDataTable) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim ObjItem As New ClsTransaksiPOItem
        Dim CekNomorSama As Boolean = False
        Dim ObjSupplier As New Core_Module.clsMasterSupplier
        Dim TipeTransaksi As Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If Me.ApakahDataBenar(False) = True Then
                'cek penomoran sama
                'If Format(Me.mTanggalTransaksi, "MMM yyyy") <> Format(Me.GetTanggalPO(Me.mPrimary), "MMM yyyy") Then
                '    Me.CekNomorSimpanPerubahan()
                'End If
                'hitung level otorisasi yang dibutuhkan
                Dim TaxExclude As Integer = 0
                If Me.TaxExclude = True Then
                    TaxExclude = 1
                End If
                If ObjSupplier.ApakahLokal(Me.mPrimarySupp) = True Then
                    TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal
                Else
                    TipeTransaksi = Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Impor
                End If
                Me.mNoUrut = Me.ObjNumbering.GetUrutan(Me.mNoTransaksi, TipeTransaksi, Me.mTanggalTransaksi)

                If DTA.Update(Me.primarySupplier, Me.NoTransaksi, Me.urutan, Me.TanggalTransaksi, _
                Me.IdentifierSuppAddress, Me.SuppContactPerson, Me.SuppPhoneNo, Me.PaymentTerm, _
                Me.ShipVia, Me.PaymentDue, Me.ReferencePaymentDue, Me.PrimaryPurchaseStaffAuthorization, _
                Me.PurchaseStaffAuthorizationDate, Me.PrimaryManagerStaffAuthorization, Me.PurchaseManagerAuthorizationDate, _
                Me.PrimaryDirectorStaffAuthorization, Me.PurchaseDirectorAuthorizationDate, Me.Authorizationlevel, _
                Me.AuthorizationNeeded, Me.PrimaryCurrency, Me.SubTotal, TaxExclude, Me.Tax, Me.TaxAmount, Me.Total, _
                Me.Remarks, vUserCode, Now, Now, Me.mPreparedBy, Me.mPreparedBy_Jabatan, Me.mApprovedBy, _
                Me.mApprovedBy_Jabatan, Me.mAcknowledgeBy, Me.mAcknowlegdeBy_Jabatan, Me.mIdentifierShipTo, _
                Me.mNeedDP, Me.mDPPercent, Me.mDPAmount, Me.mDPFromPercent, Me.primary, Me.TimeStamp) > 0 Then
                    'update data yang berubah setelah simpan perubahan
                    Me.mTimestamp = DTA.ScalarQuery_GetTimeStamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.Hapus(Me.mPrimary)
                    'simpan data item baru
                    ObjItem.Simpan(dgvItem, Me.mPrimary, vFrm, vTBLItemDelivery)
                    'set purchase order balance
                    ObjItem.SetOrderBalance(Me.mPrimary)
                    'nilai balikan
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            If ex.Message = "Duplicate entry '" + Me.mNoTransaksi + "' for key 2" Then
                'untuk yang no transaksinya sudah pernah ada, maka ubah no transaksi dan coba simpan lagi
                If Me.ObjNumbering.Apakah_Autonumber(Core_Module.ClsMasterTransactionNumbering.EnumTipeTransaksiNumbering.PO_Lokal) = True Then
                    Me.mNoTransaksi = Me.GetNoBaru1(Me.mTanggalTransaksi, Me.mPrimarySupp)
                    Return Me.SimpanPerubahan1(dgvItem, vUserCode, vFrm, vTBLItemDelivery)
                Else
                    MsgBox("PO Number duplication, please change PO number.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            End If
        Finally
            DTA.Dispose()
            ObjItem = Nothing
            ObjSupplier = Nothing
        End Try
    End Function

    ''' <summary>
    ''' fungsi ini hanya dipanggil oleh fungsi simpan perubahan, untuk mengecek jika ada yang mirip dengan nomor transksi ini. Hal ini
    ''' bisa diterjadi dimana transaksi po berubah bulan, yang kemudian menjadi sama dengan nomor transaksi lainnya. 
    ''' Misal A edit po pindah bulan, B bikin PO pada bulan ybs. B simpan duluan, saat A masih belum simpan, tapi suda berubah
    ''' tanggal, hal ini kemudian akan mengakibatkankan nomor kedua transaksi ini sama, oleh karena itu lah fungsi ini dijalankan
    ''' </summary>
    ''' <remarks></remarks>
    'Private Sub CekNomorSimpanPerubahan()
    '    Try

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Sub

    Public Overridable Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
    ByVal StrictSalesStaffDelete As Boolean, ByVal vSalesStaff As String, _
    Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        'Return True
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim CreatedBy As String = ""
        Dim DTA1 As New DseTransaksiAOPTableAdapters.tbltrans_aop1TableAdapter
        Dim TBL1 As New DseTransaksiAOP.tbltrans_aop1DataTable
        Dim row1 As DseTransaksiAOP.tbltrans_aop1Row
        Dim Dummy As String = ""
        Dim ObjUser As New Core_Module.ClsUser
        Dim DTA2 As New Core_Module.DseBalanceDPTableAdapters.tbltrans_historicalbalancemanual_dpTableAdapter
        Dim TBL2 As New Core_Module.DseBalanceDP.tbltrans_historicalbalancemanual_dpDataTable
        Dim row2 As Core_Module.DseBalanceDP.tbltrans_historicalbalancemanual_dpRow
        Dim TipeBalanceDP As Core_Module.ClsTransaksiBalanceDP.EnumTipeDP = Core_Module.ClsTransaksiBalanceDP.EnumTipeDP.DP_AOP

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Connection.ConnectionString = DTA.Connection.ConnectionString 

            'jika sudah melewati batas penguncian transaksi maka tidak bisa hapus data
            'msgbox("belum")

            'data SO boleh dihapus jika
            '1. tidak ada GRN yang merujuk pada Po ybs
            'Dim ObjGRNItem As New ClsTransaksiGRNItem
            'If ObjGRNItem.GetJmlByPO(vPrimaryTransaksi) > 0 Then
            '    If TampilkanPesan = True Then
            '        MsgBox("You cannot delete this PO because there are GRN refering to this PO", _
            '       MsgBoxStyle.Information, "Information")
            '    End If
            '    Return False
            'End If
            'ObjGRNItem = Nothing

            '1. Sudah ada stock in purchase yang merujuk pada item po ini
            Dim objStockInOut As New Inventory_Module.clsTransaksiStockInOut
            If objStockInOut.GetJmlByPrimaryPO(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this PO because there are stock in purchase refering to this PO", _
                   MsgBoxStyle.Information, "Information")
                End If
                objStockInOut = Nothing
                Return False
            End If

            objStockInOut = Nothing

            '2. tidak ada POC yang merujuk ke transaksi ybs
            Dim ObjPOC As New ClsTransaksiPOC
            If ObjPOC.GetJmlByPO(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("You cannot delete this PO because there are POC refering to this PO", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjPOC = Nothing

            '3. Jika sudah di otorisasi, maka tidak bisa diedit
            If Me.IsSkipCheckAuthrization1(vSalesStaff, ObjUser) = False Then
                If DTA.ScalarQuery_getAuthorizationLevel(vPrimaryTransaksi).HasValue = False Then
                Else
                    If DTA.ScalarQuery_getAuthorizationLevel(vPrimaryTransaksi).Value <> 0 Then
                        If TampilkanPesan = True Then
                            MsgBox("This transaction have been authorize, please cancel authorization first.", _
                            MsgBoxStyle.Information, "Information")
                        End If
                        Return False
                    Else
                    End If
                End If
            End If


            '4. Jika strict editting maka hanya sales staff yang menciptakan yang boleh edit
            CreatedBy = DTA.ScalarQuery_getCreatedBy(vPrimaryTransaksi)
            If CreatedBy = vSalesStaff Then
            Else
                If ObjUser.ApakahMemilikiHakAkses(vSalesStaff, "FrmTransaksiPO_Override_Delete") = False Then
                    If TampilkanPesan = True Then
                        MsgBox("Only user with ID:" + CreatedBy + " can delete this data.", _
                        MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If

            'po ini adalah po yang diciptakan untuk quick purchase
            If DTA.ScalarQuery_GetPrimaryAOPQuick(vPrimaryTransaksi).HasValue = True Then
                Dim ObjAOP As New ClsTransaksiAOP
                Dim PrimaryAOP As Double = DTA.ScalarQuery_GetPrimaryAOPQuick(vPrimaryTransaksi).Value
                If PrimaryAOP <> -1 Then
                    MsgBox("This PO is refered to quick purchase: " + _
                    ObjAOP.GetNoTransaksi(PrimaryAOP) + ". You cannot delete this data.", MsgBoxStyle.Information, "Information")
                    ObjAOP = Nothing
                    Return False
                End If
                ObjAOP = Nothing
            End If

            'po sudah dirujuk oleh purchase downpayment
            DTA1.Fill(TBL1, vPrimaryTransaksi)
            If TBL1.Rows.Count > 0 Then
                For Each row1 In TBL1.Rows
                    If Dummy = "" Then
                        Dummy = row1.No_Transaksi
                    Else
                        Dummy = Dummy + ", " + row1.No_Transaksi
                    End If
                Next

                MsgBox("This purchase order if refered by purchase down payment: " + Dummy + ". You cannot delete this data.", _
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

                MsgBox("This purchase order refered by down payment AOP (begining balance): " + Dummy + ". You cannot delete this data.", _
                MsgBoxStyle.Information, "Information")

                Return False
            End If

            'nilai balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjUser = Nothing
            DTA.Dispose()
            DTA1.Dispose()
            TBL1.Dispose()
            DTA2.Dispose()
            TBL2.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapusItem(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double, Optional ByVal vShowWarning As Boolean = True) As Boolean
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po4TableAdapter
        Dim NoD As New Nullable(Of Double)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            ''cek ke POC
            NoD = dta.ScalarQuery_GetJmlInPOC(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                If NoD.Value > 0 Then
                    If vShowWarning = True Then
                        MsgBox("You cannot delete this product because there's already POC refering to this product.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            ''cek ke Stock in Purchase
            NoD = dta.ScalarQuery_GetJmlInStockInPurchase(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                If NoD.Value > 0 Then
                    If vShowWarning = True Then
                        MsgBox("You cannot delete this product because there's already Stock in Purchase refering to this product.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            ''cek ke Packing List
            NoD = dta.ScalarQuery_GetJmlInPackingList(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                If NoD.Value > 0 Then
                    If vShowWarning = True Then
                        MsgBox("You cannot delete this product because there's already Packing List refering to this product.", MsgBoxStyle.Information, "Information")
                    End If
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
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
            Me.vPrimary_Antrian, vPrimaryTransaksi, Core_Module.ClsAntrian.EnumTipeTransaksiAsalAntrian.TransaksiPurchaseOrder_Hapus, UserCode)

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
                            retval = Me.Hapus1(vPrimaryTransaksi, UserCode, True)
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

    Public Overridable Function Hapus1(ByVal vPrimaryTransaksi As Double, _
    ByVal UserCode As String, ByVal StrictSalesStaffDelete As Boolean) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            '  If Me.ApakahBolehHapus(vPrimaryTransaksi, StrictSalesStaffDelete, _
            ' UserCode) = True Then
            'hapus data header, data item otomatis dihapuskan oleh server
            If DTA.Delete(vPrimaryTransaksi) > 0 Then
                Return True
            End If
            'End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetTimeStamp(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTimeStamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Authorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
      ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2Row
        Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For Each rowview In clbAutho.CheckedItems
                row = rowview.Row
                If vAuthorizationLevel = 0 Then
                    DTA.UpdateQuery_PurchaseStaffAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                    'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv2") = True Then
                        DTA.UpdateQuery_PurchaseManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1 + 1, row.Primary)
                        'otorisasi lanjutan jika punya hak otorisasi
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv3") = True Then
                            DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                            vAuthorizationLevel + 1 + 1 + 1, row.Primary)
                        End If
                    End If
                ElseIf vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_PurchaseManagerAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
                    'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv3") = True Then
                        DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1 + 1, row.Primary)
                    End If
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                    vAuthorizationLevel + 1, row.Primary)
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

    Public Function Authorize(ByVal vPrimaryPO As Double, ByVal vPrimaryUser As String, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'lvl 1
            DTA.UpdateQuery_PurchaseStaffAuthorization(vPrimaryUser, Now, _
                1, vPrimaryPO)
            'lvl 2
            DTA.UpdateQuery_PurchaseManagerAuthorization(vPrimaryUser, Now, _
                2, vPrimaryPO)
            'lvl 3
            DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
            3, vPrimaryPO)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CancelAuthorize(ByVal clbAutho As CheckedListBox, ByVal vPrimaryUser As String, _
   ByVal vAuthorizationLevel As Integer) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2Row
        Dim rowview As DataRowView
        Dim ObjUser As New Core_Module.ClsUser
        Dim vUserCode As String = ""

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vUserCode = ObjUser.GetUserCode(vPrimaryUser)

            For i As Integer = 0 To clbAutho.CheckedItems.Count - 1
                rowview = clbAutho.CheckedItems(i)
                row = rowview.Row
                If vAuthorizationLevel = 1 Then
                    DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                ElseIf vAuthorizationLevel = 2 Then
                    DTA.UpdateQuery_PurchaseManagerAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                    'cancel otorisasi lanjutan jika memang punya hak akses atas hal tersebut, 
                    'dan yang melakukan otorisasi adalah orang yang bersangkutan
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv1") = True Then
                        If Me.GetAuthorizeBy_Lvl1(row.Primary) = vPrimaryUser = -1 Then
                            DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
                            vAuthorizationLevel - 1 - 1, row.Primary)
                        End If
                    End If
                ElseIf vAuthorizationLevel = 3 Then
                    DTA.UpdateQuery_DirectorAuthorization(-1, #1/1/1950#, _
                    vAuthorizationLevel - 1, row.Primary)
                    'cancel otorisasi lanjutan jika memang punya hak akses untuk melakukan hal tersebut
                    'hanya berlaku jika orang yang bersangkutan yang melakukan otorisasi
                    If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv2") = True Then
                        If Me.GetAuthorizeBy_Lvl2(row.Primary) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl2(row.Primary) = -1 Then
                            DTA.UpdateQuery_PurchaseManagerAuthorization(-1, #1/1/1950#, _
                            vAuthorizationLevel - 1 - 1, row.Primary)

                            'lanjut jika yang mengotorisasi lvl 1 juga ybs dan ybs punya hak akses untuk cancel otorisasi
                            'pada lvl 1
                            If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv1") = True Then
                                If Me.GetAuthorizeBy_Lvl1(row.Primary) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl1(row.Primary) = -1 Then
                                    DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
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
            ObjUser = Nothing
        End Try
    End Function

    Public Function CancelAuthorize(ByVal vPrimaryPO As Double, ByVal vPrimaryUser As String, ByVal vUserCode As String) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'lv 3
            DTA.UpdateQuery_DirectorAuthorization(-1, #1/1/1950#, _
                2, vPrimaryPO)
            'lv 2
            DTA.UpdateQuery_PurchaseManagerAuthorization(-1, #1/1/1950#, _
                1, vPrimaryPO)
            'lv 1
            DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
                0, vPrimaryPO)

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiChecklistBoxBatalOtorisasi(ByVal vLevel As Integer, _
  ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiPO.tbltranspo_authorize_view2DataTable, _
  Optional ByVal StrictSalesStaff As Boolean = True, _
  Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltranspo_authorize_view2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'mengisikan semua quotation yang belum diotorisasi
            If StrictSalesStaff = True Then
                If vLevel = 1 Then
                    DTA.FillBy_UserNAuthorizationLevel_Cancel(TBL, vLevel, SalesStaff)
                ElseIf vLevel = 2 Then
                    DTA.FillBy_AuthorizationLevel_Cancel_PurchaseManager(TBL, SalesStaff)
                ElseIf vLevel = 3 Then
                    DTA.FillBy_AuthorizationLevel_Cancel_Director(TBL, SalesStaff)
                End If
            Else
                DTA.FillBy_AuthorizationLevel_Cancel(TBL, vLevel)
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

    Public Function IsiChecklistBoxOtorisasi(ByVal vLevel As Integer, _
ByVal clb As CheckedListBox, ByVal TBL As DseTransaksiPO.tbltranspo_authorize_view2DataTable, _
Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "") As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltranspo_authorize_view2TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            'mengisikan semua quotation yang belum diotorisasi
            If vLevel = 0 Then
                If StrictSalesStaff = True Then
                    DTA.FillBy_UserNAuthorizationLevel(TBL, vLevel, SalesStaff)
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
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

    Public Function GetPrimary(ByVal vNoTrans As String) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPrimary(vNoTrans).HasValue = False Then
                Return -1
            Else
                Return DTA.ScalarQuery_GetPrimary(vNoTrans).Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return -1
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_getPrimaryCurrency(vPrimary).HasValue Then
                Return DTA.ScalarQuery_getPrimaryCurrency(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPaymentDue(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetPaymentDueDate(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetPaymentDueDate(vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimarySupplier(ByVal vPrimary As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetSupplierByPrimary(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetSupplierByPrimary(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal Criteria As String, _
    ByVal SearchCriteria As EnumSearchCriteria, Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Long
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                Return DTA.ScalarQuery_GetJmlNoTransaksiLike(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                Return DTA.ScalarQuery_GetJmlTanggalDAriSampai(vTanggalDari, vTanggalSampai).Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                Return DTA.ScalarQuery_GetJmlByProductCodeLIKE(Criteria + "%").Value
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                Return DTA.ScalarQuery_GetJmlProductDescLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                Return DTA.ScalarQuery_GetJmlSpecLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                Return DTA.ScalarQuery_GetJmlBySuppCodeLike(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                Return DTA.ScalarQuery_GetJmlSuppNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.By_IPO_Number Then
                Return DTA.ScalarQuery_GetJmlSuppNameLIKE(Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemarksLIKE(Criteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOByAll(ByVal vDGVIPO As DataGridView) As Boolean

    End Function

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetInfo(ByVal vDGVIPO As DataGridView, ByVal vPrimaryIPO As String, ByVal vPrimaryProduct As String) As Boolean

    End Function

    Public Overridable Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseTransaksiPO.tbltrans_po_view0DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vTanggalDari As Date = (#1/1/1950#), _
    Optional ByVal vTanggalSampai As Date = (#1/1/1950#)) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_view0TableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByNoTransaksi Then
                QueryResult = DTA.FillBy_NoTransaksiPOLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByTanggalTransaksi Then
                QueryResult = DTA.FillBy_TgldariSampai(TBL, vTanggalDari, vTanggalSampai)
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductCodeNo Then
                QueryResult = DTA.FillBy_ProductCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductDesc Then
                QueryResult = DTA.FillBy_ProductDescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByProductSpec Then
                QueryResult = DTA.FillBy_ProductSpecLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByKodeSupplier Then
                QueryResult = DTA.FillBy_SuppCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.BySupplierName Then
                QueryResult = DTA.FillBy_SuppNameLIKE(TBL, Criteria + "%")
                'ElseIf SearchCriteria = EnumSearchCriteria.By_IPO_Number Then
                'QueryResult = DTA.FillBy_NoIPO_Like(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(2).DefaultCellStyle.Format = "dd MMM yyyy"
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

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <param name="vSupplierCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOBySupplierCode(ByVal vDGVIPO As DataGridView, ByVal vSupplierCode As String) As Boolean

    End Function

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <param name="vSupplierName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOBySupplierName(ByVal vDGVIPO As DataGridView, ByVal vSupplierName As String) As Boolean

    End Function

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOByProductCode(ByVal vDGVIPO As DataGridView, ByVal vProductCode As String) As Boolean

    End Function

    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOByProductDesc(ByVal vDGVIPO As DataGridView, ByVal vDesc As String) As Boolean

    End Function
    ''' <summary>
    ''' Fungsi ini digunakan pada transaksi po untuk mengambil data ipo yang belum dimiliki oleh po
    ''' </summary>
    ''' <param name="vDGVIPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsiDGVIPOByIPONo(ByVal vDGVIPO As DataGridView, ByVal vNoIPO As String) As Boolean

    End Function

    ''' <summary>
    ''' mengambil item dari tabel dan mengumpulkannya berdasarkan category product yang terjadi
    ''' </summary>
    ''' <param name="vTBL"></param>
    ''' <remarks></remarks>
    Private Sub SetSizing(ByVal vTBL As DseTransaksiPO.TblCetakDataTable)
        Dim TBLDummy As New DseTransaksiPO.TblCetakDataTable
        Dim row, row1 As DseTransaksiPO.TblCetakRow
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
                                row1.Qty_36 += row.Qty
                            ElseIf row.Size = 37 Then
                                row1.Qty_37 += row.Qty
                            ElseIf row.Size = 38 Then
                                row1.Qty_38 += row.Qty
                            ElseIf row.Size = 39 Then
                                row1.Qty_39 += row.Qty
                            ElseIf row.Size = 40 Then
                                row1.Qty_40 += row.Qty
                            ElseIf row.Size = 41 Then
                                row1.Qty_41 += row.Qty
                            ElseIf row.Size = 42 Then
                                row1.Qty_42 += row.Qty
                            ElseIf row.Size = 43 Then
                                row1.Qty_43 += row.Qty
                            ElseIf row.Size = 44 Then
                                row1.Qty_44 += row.Qty
                            ElseIf row.Size = 45 Then
                                row1.Qty_45 += row.Qty
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

                            If row.Size = 36 Then
                                .Qty_36 = row.Qty
                            ElseIf row.Size = 37 Then
                                .Qty_37 = row.Qty
                            ElseIf row.Size = 38 Then
                                .Qty_38 = row.Qty
                            ElseIf row.Size = 39 Then
                                .Qty_39 = row.Qty
                            ElseIf row.Size = 40 Then
                                .Qty_40 = row.Qty
                            ElseIf row.Size = 41 Then
                                row1.Qty_41 += row.Qty
                            ElseIf row.Size = 42 Then
                                row1.Qty_42 += row.Qty
                            ElseIf row.Size = 43 Then
                                row1.Qty_43 += row.Qty
                            ElseIf row.Size = 44 Then
                                row1.Qty_44 += row.Qty
                            ElseIf row.Size = 45 Then
                                row1.Qty_45 += row.Qty
                            End If

                            .Amount = .Nett_Price * .Qty
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

    Private Function CetakPO3Banyak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal ListPrimaryPo As List(Of Double)) As Boolean
        Dim DS As New DseTransaksiPO
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim tblDummy As New DseTransaksiPO.TblCetakDataTable
        Dim rowDummy As DseTransaksiPO.TblCetakRow
        Dim tblDummyP As New DseTransaksiPO.TblGetIsiParameterCetakDataTable
        Dim rowDummyP As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 0
        Dim row1 As DseTransaksiPO.TblCetakRow
        Dim row2 As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim DTA1 As New DseTransaksiPOTableAdapters.TblCetakTableAdapter
        Dim DTA2 As New DseTransaksiPOTableAdapters.TblGetIsiParameterCetakTableAdapter
        Dim Add As String = ""
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow

        Try
            DS.TblCetak.Rows.Clear()
            DS.TblGetIsiParameterCetak.Rows.Clear()
            'DTA1.FillBy_PrimaryPO(DS.TblCetak, vPrimaryPO)
            DTA1.Fill(tblDummy)
            DTA2.FillBy(tblDummyP)
            DTA1.Dispose()
            Dim strFilter As String = ""
            strFilter = " [Primary PO] = [Primary PO] "

            If ListPrimaryPo.Count > 0 Then
                Dim strFilter2 As String = ""
                For intA As Integer = 0 To ListPrimaryPo.Count - 1
                    If strFilter2 = "" Then
                        strFilter2 &= " [Primary PO] = " & ListPrimaryPo.Item(intA)
                    Else
                        strFilter2 &= " Or [Primary PO] = " & ListPrimaryPo.Item(intA)
                    End If
                Next
                If strFilter2 <> "" Then
                    strFilter &= " And (" & strFilter2 & ")"
                End If
            End If


            tblDummy.DefaultView.RowFilter = Nothing
            tblDummy.DefaultView.RowFilter = strFilter
            tblDummy.DefaultView.Sort = "[Primary PO]"
            If tblDummy.DefaultView.Count > 0 Then
                For intA As Integer = 0 To tblDummy.DefaultView.Count - 1
                    DS.TblCetak.Rows.Add()
                    For intC As Integer = 0 To tblDummy.Columns.Count - 1
                        DS.TblCetak.Rows(intA)(intC) = tblDummy.DefaultView.Item(intA)(intC)
                    Next
                Next
            End If


            tblDummyP.DefaultView.RowFilter = Nothing
            tblDummyP.DefaultView.RowFilter = strFilter
            tblDummyP.DefaultView.Sort = "[Primary PO]"
            If tblDummyP.DefaultView.Count > 0 Then
                For intA As Integer = 0 To tblDummyP.DefaultView.Count - 1
                    DS.TblGetIsiParameterCetak.Rows.Add()
                    For intC As Integer = 0 To tblDummyP.Columns.Count - 1
                        DS.TblGetIsiParameterCetak.Rows(intA)(intC) = tblDummyP.DefaultView.Item(intA)(intC)
                    Next
                Next
            End If



            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)
            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                'regroup berdasarkan sizing
                Me.SetSizing(DS.TblCetak)
                row = DS.TblGetIsiParameterCetak(0)
                If row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str <> "" Then
                    Dim Report As New CrpCetakTransaksiPO3
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
                ElseIf row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str = "" Then
                    Dim Report As New CrpCetakTransaksiPO3_2posisi
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
                Else
                    Dim Report As New CrpCetakTransaksiPO3_1posisi

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
            End If
            DTA2.Dispose()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Private Function CetakPO3(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPO As Double) As Boolean
        Dim DS As New DseTransaksiPO
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = ObjCur.GetDefaultPrecission(Me.GetPrimaryCurrency(vPrimaryPO))
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim DTA1 As New DseTransaksiPOTableAdapters.TblCetakTableAdapter
        Dim DTA2 As New DseTransaksiPOTableAdapters.TblGetIsiParameterCetakTableAdapter
        Dim Add As String = ""

        Try
            DTA1.FillBy_PrimaryPO(DS.TblCetak, vPrimaryPO)
            DTA1.Dispose()

            DTA2.Fill(DS.TblGetIsiParameterCetak, vPrimaryPO)
            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)
            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                'regroup berdasarkan sizing
                Me.SetSizing(DS.TblCetak)

                row = DS.TblGetIsiParameterCetak(0)
                If row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str <> "" Then
                    Dim Report As New CrpCetakTransaksiPO3

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
                ElseIf row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str = "" Then
                    Dim Report As New CrpCetakTransaksiPO3_2posisi

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
                Else
                    Dim Report As New CrpCetakTransaksiPO3_1posisi

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
            End If
            DTA2.Dispose()

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
        End Try
    End Function

    Public Function GetPOSubTotal(ByVal vPrimaryPO As Double) As Decimal
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetSubTotal(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetSubTotal(vPrimaryPO).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function GetAddress(ByVal vTBL As DseTransaksiPO.TblGetIsiParameterCetakDataTable) As String
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim dummy As String = ""

        Try
            If vTBL.Rows.Count > 0 Then
                row = vTBL.Rows(0)
                If row.Address1 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address1
                    Else
                        dummy = dummy + vbCrLf + row.Address1
                    End If
                End If
                If row.Address2 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address2
                    Else
                        dummy = dummy + vbCrLf + row.Address2
                    End If
                End If
                If row.Address3 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address3
                    Else
                        dummy = dummy + vbCrLf + row.Address3
                    End If
                End If
                If row.Address4 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address4
                    Else
                        dummy = dummy + vbCrLf + row.Address4
                    End If
                End If
                If row.Address5 <> "" Then
                    If dummy = "" Then
                        dummy = row.Address5
                    Else
                        dummy = dummy + vbCrLf + row.Address5
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

    ''' <summary>
    ''' mengembalikan nilai true untuk po yang memiliki item yang menggunakan sizing
    ''' </summary>
    ''' <param name="vPrimaryPO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function UseSizing(ByVal vPrimaryPO As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryPO).HasValue = True Then
                If DTA.ScalarQuery_GetJmlUseSizing(vPrimaryPO).Value > 0 Then
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

    Public Function CetakPOBanyak(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal ListPrimaryPO As List(Of Double)) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objPO_Custom As New CustomSMESS.ClsTransaksiPO
        'Dim frmPO_Custom As New CustomSMESS.FrmTransaksiPO_Print
        'Dim frmPO_Custom_Opt As New CustomSMESS.FrmTransaksiPO_PrintDialog

        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            ObjSetting = Nothing
            Return objPO_Custom.CetakPOBanyak(CRV, ListPrimaryPO)
            'frmPO_Custom_Opt.ShowDialog()
            'If frmPO_Custom_Opt.CboFormat.SelectedIndex = 0 Then
            '    frmPO_Custom.mBentukCetakan = CustomSMESS.FrmTransaksiPO_Print.EnumBentukCetakan.F4
            'Else
            '    frmPO_Custom.mBentukCetakan = CustomSMESS.FrmTransaksiPO_Print.EnumBentukCetakan.F4_Half
            'End If
            'frmPO_Custom.mPrimaryPO = vPrimaryPO
            'frmPO_Custom.ShowDialog()
            'Exit Function
        End If

        ''If Me.UseSizing(vPrimaryPO) = True Then
        ''    Return Me.CetakPO3Banyak(CRV, ListPrimaryPO)
        ''    Exit Function
        ''End If


        Dim DS As New DseTransaksiPO
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        'Dim FormatDecimal As Integer = ObjCur.GetDefaultPrecission(Me.GetPrimaryCurrency(vPrimaryPO))
        Dim FormatDecimal As Integer = 0
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim Dummy As String = ""
        Dim tblDummy As New DseTransaksiPO.TblCetakDataTable
        Dim rowDummy As DseTransaksiPO.TblCetakRow
        Dim rowCetak As DseTransaksiPO.TblCetakRow
        Dim tblDummyP As New DseTransaksiPO.TblGetIsiParameterCetakDataTable
        Dim rowDummyP As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim rowParam As DseTransaksiPO.TblGetIsiParameterCetakRow


        Try
            Dim DTA1 As New DseTransaksiPOTableAdapters.TblCetakTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.Fill(tblDummy)
            DTA1.Dispose()

            Dim DTA2 As New DseTransaksiPOTableAdapters.TblGetIsiParameterCetakTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.FillBy(tblDummyP)



            Dim Filter As String = ""
            Dim Filter2 As String = ""
            Filter = " [Primary Transaksi] = [Primary Transaksi] "
            Filter2 = " [Primary PO] = [Primary PO] "

            If ListPrimaryPO.Count > 0 Then
                Dim str As String = ""
                Dim str2 As String = ""
                For intA As Integer = 0 To ListPrimaryPO.Count - 1
                    If str = "" Then
                        str &= " [Primary Transaksi] = " & ListPrimaryPO.Item(intA)
                    Else
                        str &= " OR [Primary Transaksi] = " & ListPrimaryPO.Item(intA)
                    End If

                    If str2 = "" Then
                        str2 &= " [Primary PO] = " & ListPrimaryPO.Item(intA)
                    Else
                        str2 &= " OR [Primary PO] = " & ListPrimaryPO.Item(intA)
                    End If
                Next
                Filter &= "And (" & str & ")"

                Filter2 &= "And (" & str2 & ")"
            End If




            tblDummy.DefaultView.RowFilter = Nothing
            tblDummy.DefaultView.RowFilter = Filter
            tblDummy.DefaultView.Sort = "[Primary Transaksi]"
            While tblDummy.DefaultView.Count > 0
                rowDummy = tblDummy.DefaultView(0).Row
                rowCetak = DS.TblCetak.NewTblCetakRow()
                For ijk As Integer = 0 To tblDummy.Columns.Count - 1
                    rowCetak.Item(ijk) = rowDummy.Item(ijk)
                Next
                DS.TblCetak.Rows.Add(rowCetak)
                tblDummy.Rows.Remove(rowDummy)
            End While



            tblDummyP.DefaultView.RowFilter = Nothing
            tblDummyP.DefaultView.RowFilter = Filter2
            tblDummyP.DefaultView.Sort = "[Primary PO]"
            While tblDummyP.DefaultView.Count > 0
                rowDummyP = tblDummyP.DefaultView(0).Row
                rowParam = DS.TblGetIsiParameterCetak.NewRow()
                For ijk As Integer = 0 To tblDummyP.Columns.Count - 1
                    rowParam.Item(ijk) = rowDummyP.Item(ijk)
                Next
                DS.TblGetIsiParameterCetak.Rows.Add(rowParam)
                tblDummyP.Rows.Remove(rowDummyP)
            End While



            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)
            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                row = DS.TblGetIsiParameterCetak(0)
                If row.Address1 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address1
                    Else
                        Dummy = Dummy + vbCrLf + row.Address1
                    End If
                End If
                If row.Address2 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address2
                    Else
                        Dummy = Dummy + vbCrLf + row.Address2
                    End If
                End If
                If row.Address3 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address3
                    Else
                        Dummy = Dummy + vbCrLf + row.Address3
                    End If
                End If
                If row.Address4 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address4
                    Else
                        Dummy = Dummy + vbCrLf + row.Address4
                    End If
                End If
                If row.Address5 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address5
                    Else
                        Dummy = Dummy + vbCrLf + row.Address5
                    End If
                End If
                row.Contact_Person = row.Contact_Person.Trim
                If row.Contact_Person <> "" Then
                    If Dummy = "" Then
                        Dummy = "Attn : " + row.Contact_Person
                    Else
                        Dummy = Dummy + vbCrLf + "Attn : " + row.Contact_Person
                    End If
                End If


                Dim primaryMax As Double = -1
                DS.TblCetak.DefaultView.RowFilter = Nothing
                DS.TblCetak.DefaultView.Sort = "[Primary Transaksi]"
                If DS.TblCetak.DefaultView.Count > 0 Then
                    rowCetak = DS.TblCetak.DefaultView(DS.TblCetak.DefaultView.Count - 1).Row
                    primaryMax = rowCetak.Primary_Transaksi
                End If
                DS.TblCetak.DefaultView.RowFilter = Nothing


                If ListPrimaryPO.Count > 0 Then
                    If row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str <> "" Then
                        Dim Report As New CrpCetakTransaksiPO_Banyak

                        Report.SetDataSource(DS)

                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)
                        Report.SetParameterValue(Report.Parameter_MaxPrimary.ParameterFieldName, primaryMax)

                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    ElseIf row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str = "" Then
                        Dim Report As New CrpCetakTransaksiPO_2posisi_Banyak
                        Report.SetDataSource(DS)
                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)
                        Report.SetParameterValue(Report.Parameter_MaxPrimary.ParameterFieldName, primaryMax)
                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    Else
                        Dim Report As New CrpCetakTransaksiPO_1posisi_Banyak
                        Report.SetDataSource(DS)
                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)
                        Report.SetParameterValue(Report.Parameter_MaxPrimary.ParameterFieldName, primaryMax)
                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    End If
                Else
                    If row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str <> "" Then
                        Dim Report As New CrpCetakTransaksiPO

                        Report.SetDataSource(DS)

                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    ElseIf row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str = "" Then
                        Dim Report As New CrpCetakTransaksiPO_2posisi

                        Report.SetDataSource(DS)

                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    Else
                        Dim Report As New CrpCetakTransaksiPO_1posisi

                        Report.SetDataSource(DS)

                        Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                        Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                        Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                        Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                        Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                        Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                        Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                        CRV.ReportSource = Report
                        CRV.Show()
                        Return True
                    End If
                End If

            End If
            DTA2.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            objPO_Custom = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Public Function CetakPO(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer, ByVal vPrimaryPO As Double) As Boolean
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim objPO_Custom As New CustomSMESS.ClsTransaksiPO
        'Dim frmPO_Custom As New CustomSMESS.FrmTransaksiPO_Print
        'Dim frmPO_Custom_Opt As New CustomSMESS.FrmTransaksiPO_PrintDialog

        If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.PO, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Use_Custom_Print_Logic) = "1" Then
            ObjSetting = Nothing
            Return objPO_Custom.CetakPO(CRV, vPrimaryPO)
            'frmPO_Custom_Opt.ShowDialog()
            'If frmPO_Custom_Opt.CboFormat.SelectedIndex = 0 Then
            '    frmPO_Custom.mBentukCetakan = CustomSMESS.FrmTransaksiPO_Print.EnumBentukCetakan.F4
            'Else
            '    frmPO_Custom.mBentukCetakan = CustomSMESS.FrmTransaksiPO_Print.EnumBentukCetakan.F4_Half
            'End If
            'frmPO_Custom.mPrimaryPO = vPrimaryPO
            'frmPO_Custom.ShowDialog()
            'Exit Function
        End If

        If Me.UseSizing(vPrimaryPO) = True Then
            Return Me.CetakPO3(CRV, vPrimaryPO)
            Exit Function
        End If

        Dim DS As New DseTransaksiPO
        Dim ObjCur As New Core_Module.clsMasterCurrency

        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = ObjCur.GetDefaultPrecission(Me.GetPrimaryCurrency(vPrimaryPO))
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim Dummy As String = ""

        Try
            Dim DTA1 As New DseTransaksiPOTableAdapters.TblCetakTableAdapter
            DTA1.Connection.ConnectionString = Core_Module.GetConstring
            DTA1.FillBy_PrimaryPO(DS.TblCetak, vPrimaryPO)
            DTA1.Dispose()
            Dim DTA2 As New DseTransaksiPOTableAdapters.TblGetIsiParameterCetakTableAdapter
            DTA2.Connection.ConnectionString = Core_Module.GetConstring
            DTA2.Fill(DS.TblGetIsiParameterCetak, vPrimaryPO)
            Dim Jumlah As Integer = DS.TblGetIsiParameterCetak.Rows.Count
            Me.Set_TandaTangan(DS.TblGetIsiParameterCetak)
            If DS.TblGetIsiParameterCetak.Rows.Count > 0 Then
                row = DS.TblGetIsiParameterCetak(0)
                If row.Address1 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address1
                    Else
                        Dummy = Dummy + vbCrLf + row.Address1
                    End If
                End If
                If row.Address2 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address2
                    Else
                        Dummy = Dummy + vbCrLf + row.Address2
                    End If
                End If
                If row.Address3 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address3
                    Else
                        Dummy = Dummy + vbCrLf + row.Address3
                    End If
                End If
                If row.Address4 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address4
                    Else
                        Dummy = Dummy + vbCrLf + row.Address4
                    End If
                End If
                If row.Address5 <> "" Then
                    If Dummy = "" Then
                        Dummy = row.Address5
                    Else
                        Dummy = Dummy + vbCrLf + row.Address5
                    End If
                End If
                row.Contact_Person = row.Contact_Person.Trim
                If row.Contact_Person <> "" Then
                    If Dummy = "" Then
                        Dummy = "Attn : " + row.Contact_Person
                    Else
                        Dummy = Dummy + vbCrLf + "Attn : " + row.Contact_Person
                    End If
                End If

                If row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str <> "" Then
                    Dim Report As New CrpCetakTransaksiPO

                    Report.SetDataSource(DS)

                    Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                    Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                    Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                    Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                    Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                    Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                    Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                    CRV.ReportSource = Report
                    CRV.Show()
                    Return True
                ElseIf row.Prepared_By_Str <> "" And row.Approved_By_Str <> "" And row.Acknowledge_By_Str = "" Then
                    Dim Report As New CrpCetakTransaksiPO_2posisi

                    Report.SetDataSource(DS)

                    Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                    Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                    Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                    Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                    Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                    Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                    Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                    CRV.ReportSource = Report
                    CRV.Show()
                    Return True
                Else
                    Dim Report As New CrpCetakTransaksiPO_1posisi

                    Report.SetDataSource(DS)

                    Report.SetParameterValue(Report.Parameter_VAT.ParameterFieldName, "VAT")
                    Report.SetParameterValue(Report.Parameter_Total.ParameterFieldName, "Total")
                    Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
                    Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)
                    Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
                    Report.SetParameterValue(Report.Parameter_FormatYears.ParameterFieldName, FormatYears)
                    Report.SetParameterValue(Report.Parameter_Address.ParameterFieldName, Dummy)

                    CRV.ReportSource = Report
                    CRV.Show()
                    Return True
                End If
            End If
            DTA2.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            objPO_Custom = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    Private Sub Set_TandaTangan(ByVal TBL As DseTransaksiPO.TblGetIsiParameterCetakDataTable)
        Dim row As DseTransaksiPO.TblGetIsiParameterCetakRow
        Dim Posisi1, Posisi2, Posisi3, Nama1, Nama2, Nama3, Jabatan1, Jabatan2, Jabatan3 As String

        Try
            Posisi1 = ""
            Posisi2 = ""
            Posisi3 = ""
            Nama1 = ""
            Nama2 = ""
            Nama3 = ""
            Jabatan1 = ""
            Jabatan2 = ""
            Jabatan3 = ""

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                If row.PreparedByName <> "" Then
                    Posisi1 = "Prepared By"
                    Nama1 = row.PreparedByName
                    Jabatan1 = row.PreparedByPosition
                    If row.ApprovedByName <> "" Then
                        Posisi2 = "Approved By"
                        Nama2 = row.ApprovedByName
                        Jabatan2 = row.ApprovedByPosition

                        If row.AcknowledgeByName <> "" Then
                            Posisi3 = "Acknowledge By"
                            Nama3 = row.AcknowledgeByName
                            Jabatan3 = row.AcknowledgeByPostion
                        End If
                    Else
                        If row.AcknowledgeByName <> "" Then
                            If Posisi2 = "" Then
                                Posisi2 = "Acknowledge By"
                                Nama2 = row.AcknowledgeByName
                                Jabatan2 = row.AcknowledgeByPostion
                            Else
                                Posisi3 = "Acknowledge By"
                                Nama3 = row.AcknowledgeByName
                                Jabatan3 = row.AcknowledgeByPostion
                            End If
                        End If
                    End If
                ElseIf row.ApprovedByName <> "" Then
                    Posisi1 = "Approved By"
                    Nama1 = row.ApprovedByName
                    Jabatan1 = row.ApprovedByPosition
                    If row.AcknowledgeByName <> "" Then
                        Posisi2 = "Acknowledge By"
                        Nama2 = row.AcknowledgeByName
                        Jabatan2 = row.AcknowledgeByPostion
                    End If
                ElseIf row.AcknowledgeByName <> "" Then
                    Posisi1 = "Acknowledge By"
                    Nama1 = row.AcknowledgeByName
                    Jabatan1 = row.AcknowledgeByPostion
                End If

                For Each row In TBL.Rows
                    row.Prepared_By_Str = Posisi1
                    row.PreparedByName = Nama1
                    row.PreparedByPosition = Jabatan1
                    row.Approved_By_Str = Posisi2
                    row.ApprovedByName = Nama2
                    row.ApprovedByPosition = Jabatan2
                    row.Acknowledge_By_Str = Posisi3
                    row.AcknowledgeByName = Nama3
                    row.AcknowledgeByPostion = Jabatan3
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTaxPercentage(ByVal vPrimaryPO As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Return DTA.ScalarQuery_GetTaxPercentage(vPrimaryPO).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUseTax(ByVal vPrimaryPO As Double) As Boolean
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetUseTax(vPrimaryPO).HasValue = True Then
                Return DTA.ScalarQuery_GetUseTax(vPrimaryPO).Value
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiPO_Quick(ByVal vprimaryAOP_Quick As Double) As String
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If String.IsNullOrEmpty(DTA.ScalarQuery_GetNoPOByPrimaryAAOPQuick(vprimaryAOP_Quick)) = True Then
                Return ""
            Else
                Return DTA.ScalarQuery_GetNoPOByPrimaryAAOPQuick(vprimaryAOP_Quick)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetNoTransaksiPO(ByVal vprimary As Double) As String
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
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

    Public Function GetPrimaryCurrency_Tax(ByVal vPrimaryPO As Double, ByRef Tax As Double, _
    ByRef PrimaryCurrency As Double) As Double
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_poDataTable
        Dim row As DseTransaksiPO.tbltrans_poRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.FillBy_Primary(TBL, vPrimaryPO) > 0 Then
                row = TBL.Rows(0)

                Tax = row.Tax
                PrimaryCurrency = row.Primary_Currency
            Else
                Tax = 0
                PrimaryCurrency = -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function


    ''' <summary>
    ''' Untuk Semua No PO yang di gunakan oleh POC tertentu
    ''' </summary>
    ''' <param name="vPrimaryPOC "></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNoPOByPrimaryPOC(ByVal vPrimaryPOC As Double) As String
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
        Dim TBL As New DseTransaksiPO.tbltrans_poDataTable
        Dim row As DseTransaksiPO.tbltrans_poRow
        Dim KumpulanNoPO As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            DTA.FillBy_PrimaryPOC(TBL, vPrimaryPOC)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If KumpulanNoPO = "" Then
                        KumpulanNoPO = row.No_Transaksi
                    Else
                        KumpulanNoPO = KumpulanNoPO + ", " + row.No_Transaksi
                    End If
                Next
            End If

            Return KumpulanNoPO
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetPrimaryProduct(ByVal VPrimary As Double) As Double
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po_itemTableAdapter
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            If dta.ScalarQuery_GetPrimaryProduct(VPrimary).HasValue Then
                Return dta.ScalarQuery_GetPrimaryProduct(VPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            dta.Dispose()
        End Try
    End Function

    ''@rosadinnn
    ''@MUaT
    Public Function IsMultiUnitProduct(ByVal vPrimaryProduct As Double) As Boolean
        Dim dta As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
        Dim count_ As Nullable(Of Long) = 0
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            count_ = dta.ScalarQuery_GetCountByPrimaryProduct(vPrimaryProduct)
            If count_.HasValue = True Then
                If count_.Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function IsiCboUnit(ByVal vPrimaryProduct As Double, ByVal vCbo As ComboBox, _
            ByVal vTbl As DseTransaksiPO.CboUnitDataTable, _
            ByVal vTblMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitDataTable, _
            ByVal vTxt As TextBox) As Boolean
        Dim rowTbl As DseTransaksiPO.CboUnitRow
        Dim rowMultiUnit As DseTransaksiPO.tblmaster_product_item_multiunitRow
        Dim objProduct As New Core_Module.clsMasterProduct
        Dim objUnit As New Core_Module.ClsMasterUnit
        Try
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()
            vTbl.Rows.Clear()
            vTblMultiUnit.DefaultView.RowFilter = ""
            vTblMultiUnit.DefaultView.RowFilter = "[Primary Product] = " & vPrimaryProduct.ToString
            If vTblMultiUnit.DefaultView.Count > 0 Then
                rowTbl = vTbl.NewCboUnitRow
                rowTbl.Primary = objProduct.GetPrimaryUnit(vPrimaryProduct)
                rowTbl.Code = objUnit.GetUnitCode(rowTbl.Primary)
                vTbl.AddCboUnitRow(rowTbl)
                For i As Integer = 0 To vTblMultiUnit.DefaultView.Count - 1
                    rowMultiUnit = vTblMultiUnit.DefaultView(i).Row
                    rowTbl = vTbl.NewCboUnitRow
                    rowTbl.Primary = rowMultiUnit.Primary_Unit
                    rowTbl.Code = objUnit.GetUnitCode(rowMultiUnit.Primary_Unit)
                    vTbl.AddCboUnitRow(rowTbl)
                Next
                vCbo.DataSource = vTbl.DefaultView
                vTbl.DefaultView.Sort = "[Code] ASC"
                vCbo.DisplayMember = "Code"
                vTxt.Text = objProduct.GetPrimaryUnit(vPrimaryProduct)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return False
        End Try
    End Function

    Public Function GetRatioDefault(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
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
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetRatioReference(ByVal vPrimaryProduct As Double, ByVal vPrimaryUnit As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tblmaster_product_item_multiunitTableAdapter
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
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTotalPOC(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po4TableAdapter
        Dim NoD As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            NoD = dta.ScalarQuery_GetTotalInPOC(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                Return NoD.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTotalPackingList(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po4TableAdapter
        Dim NoD As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            NoD = dta.ScalarQuery_GetTotalInPackingList(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                Return NoD.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetTotalStockInPurchase(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po4TableAdapter
        Dim NoD As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            NoD = dta.ScalarQuery_GetTotalInStockInPurchase(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                Return NoD.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function

    Public Function GetToleranceUp(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Double
        Dim Dta As New DseTransaksiPOTableAdapters.tbltrans_po_item6TableAdapter
        Dim NoD As Nullable(Of Decimal)
        Try
            Dta.Connection.ConnectionString = Core_Module.GetConstring()
            NoD = Dta.ScalarQuery_GetToleranceUp(vPrimaryPO, vPrimaryProduct)
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

    Public Sub New()
        Core_Module.SetConstring_DTAAntrian(Me.DTAInOUT)
    End Sub

    'aul
    Public Function Authorize2(ByVal Dgv As DataGridView, ByVal vPrimaryUser As String, _
      ByVal vAuthorizationLevel As Integer) As Boolean
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
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
                        DTA.UpdateQuery_PurchaseStaffAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1, Primary_)
                        'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv2") = True Then
                            DTA.UpdateQuery_PurchaseManagerAuthorization(vPrimaryUser, Now, _
                            vAuthorizationLevel + 1 + 1, Primary_)
                            'otorisasi lanjutan jika punya hak otorisasi
                            If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv3") = True Then
                                DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                                vAuthorizationLevel + 1 + 1 + 1, Primary_)
                            End If
                        End If
                    ElseIf vAuthorizationLevel = 1 Then
                        DTA.UpdateQuery_PurchaseManagerAuthorization(vPrimaryUser, Now, _
                        vAuthorizationLevel + 1, Primary_)
                        'jika punya hak akses untuk otorisasi lanjutannya maka bisa langsung dilakukan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv3") = True Then
                            DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
                            vAuthorizationLevel + 1 + 1, Primary_)
                        End If
                    ElseIf vAuthorizationLevel = 2 Then
                        DTA.UpdateQuery_DirectorAuthorization(vPrimaryUser, Now, _
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

    Public Function CancelAuthorize2(ByVal Dgv As DataGridView, ByVal vPrimaryUser As String, _
 ByVal vAuthorizationLevel As Integer) As Boolean
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTA As New DseTransaksiPOTableAdapters.tbltrans_poTableAdapter
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
                        DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                    ElseIf vAuthorizationLevel = 2 Then
                        DTA.UpdateQuery_PurchaseManagerAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                        'cancel otorisasi lanjutan jika memang punya hak akses atas hal tersebut, 
                        'dan yang melakukan otorisasi adalah orang yang bersangkutan
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv1") = True Then
                            If Me.GetAuthorizeBy_Lvl1(Primary_) = vPrimaryUser = -1 Then
                                DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
                                vAuthorizationLevel - 1 - 1, Primary_)
                            End If
                        End If
                    ElseIf vAuthorizationLevel = 3 Then
                        DTA.UpdateQuery_DirectorAuthorization(-1, #1/1/1950#, _
                        vAuthorizationLevel - 1, Primary_)
                        'cancel otorisasi lanjutan jika memang punya hak akses untuk melakukan hal tersebut
                        'hanya berlaku jika orang yang bersangkutan yang melakukan otorisasi
                        If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv2") = True Then
                            If Me.GetAuthorizeBy_Lvl2(Primary_) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl2(Primary_) = -1 Then
                                DTA.UpdateQuery_PurchaseManagerAuthorization(-1, #1/1/1950#, _
                                vAuthorizationLevel - 1 - 1, Primary_)

                                'lanjut jika yang mengotorisasi lvl 1 juga ybs dan ybs punya hak akses untuk cancel otorisasi
                                'pada lvl 1
                                If ObjUser.ApakahMemilikiHakAkses(vUserCode, "FrmTransaksiPOOtorisasiLv1") = True Then
                                    If Me.GetAuthorizeBy_Lvl1(Primary_) = vPrimaryUser Or Me.GetAuthorizeBy_Lvl1(Primary_) = -1 Then
                                        DTA.UpdateQuery_PurchaseStaffAuthorization(-1, #1/1/1950#, _
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

    Public Function SaveUserItem(ByVal vUserCode As String, ByVal vDgvUser As DataGridView) As Boolean
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po_item_userTableAdapter
        Dim tbl As New DseTransaksiPO.tbltrans_po_item_userDataTable
        Dim vFrm As New FrmTransaksiPOOtorisasi2
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
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po_item_userTableAdapter
        Dim tbl As New DseTransaksiPO.tbltrans_po_item_userDataTable
        Dim row As DseTransaksiPO.tbltrans_po_item_userRow
        Dim vFrm As New FrmTransaksiPOOtorisasi2
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

    'aul
    Public Function IsiDgvBoxOtorisasi(ByVal vLevel As Integer, _
     ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiPO.tbltranspo_authorize_view2_2DataTable, _
     ByVal vDgvUser As DataGridView, _
     Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
     Optional ByVal vPOno As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
     Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vPoDate As Boolean = False) As Boolean
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTA As New DseTransaksiPOTableAdapters.tbltranspo_authorize_view2_2TableAdapter
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'mengisikan semua quotation yang belum diotorisasi
            If vPoDate = False Then
                If vLevel = 0 Then
                    If StrictSalesStaff = True Then
                        DTA.FillBy_UserNAuthorizationLevel(TBL, vLevel, SalesStaff)
                    Else
                        DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevel(TBL, vLevel)
                End If
            Else
                'filter dengan tanggal
                If vLevel = 0 Then
                    If StrictSalesStaff = True Then
                        DTA.FillBy_UserNAuthorizationLevel_TanggalSampaiDari(TBL, vLevel, SalesStaff, vDateFrom, vDateTo)
                    Else
                        DTA.FillBy_AuthorizationLevelTglSampaiDari(TBL, vLevel, vDateFrom, vDateTo)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevelTglSampaiDari(TBL, vLevel, vDateFrom, vDateTo)
                End If
            End If

            If vPOno <> "" Then
                filter = "[No Transaksi] = '" & vPOno.ToString & "'"
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

            If vPOno <> "" And filter2 <> "" Then
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
                    rowDGV.Cells(vFrm.ColDgvAuth_PONo.Index).Value = row.No_Transaksi
                    rowDGV.Cells(vFrm.ColDgvAuth_SupplierName.Index).Value = row.Supplier_Name
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

    'aul
    Public Function IsiDgvBoxAlreadyAuthorize(ByVal vLevel As Integer, _
     ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiPO.tbltranspo_authorize_view2_2DataTable, _
     ByVal vDgvUser As DataGridView, _
     Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
     Optional ByVal vPOno As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
     Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vStatusPO As Integer = EnumStatusPO.AllPO, _
     Optional ByVal vPoDate As Boolean = False, Optional ByVal vStatusPrint As Integer = EnumStatusPrintPO.AllPO) As Boolean
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTA As New DseTransaksiPOTableAdapters.tbltranspo_authorize_view2_2TableAdapter
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView
        Dim StatusAuth As String = ""
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            If vPoDate = False Then
                If vStatusPO = EnumStatusPO.AllPO Then
                    DTA.Fill(TBL)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_1 Then
                    DTA.FillBy_AuthorizationLevel(TBL, 1)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_2 Then
                    DTA.FillBy_AuthorizationLevel(TBL, 2)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_3 Then
                    DTA.FillBy_AuthorizationLevelDirector(TBL, 3)
                ElseIf vStatusPO = EnumStatusPO.UnauthorizedOnly Then
                    DTA.FillBy_AuthorizationLevel(TBL, 0)
                End If
            Else
                If vStatusPO = EnumStatusPO.AllPO Then
                    DTA.Fill_TanggalSampaiDari(TBL, vDateFrom, vDateTo)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_1 Then
                    DTA.FillBy_AuthorizationLevelTglSampaiDari(TBL, 1, vDateFrom, vDateTo)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_2 Then
                    DTA.FillBy_AuthorizationLevelTglSampaiDari(TBL, 2, vDateFrom, vDateTo)
                ElseIf vStatusPO = EnumStatusPO.AuthorizeLvl_3 Then
                    DTA.FillBy_AuthorizationLevelDirector_TglSampaiDari(TBL, 3, vDateFrom, vDateTo)
                ElseIf vStatusPO = EnumStatusPO.UnauthorizedOnly Then
                    DTA.FillBy_AuthorizationLevelTglSampaiDari(TBL, 0, vDateFrom, vDateTo)
                End If
            End If

            filter = "[No Transaksi] = [No Transaksi] "


            If vPOno <> "" Then
                filter &= " And [No Transaksi] = '" & vPOno.ToString & "'"
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
                filter &= " AND (" & filter2 & ")"
            End If

            If vStatusPrint = EnumStatusPrintPO.PrintedPOOnly Then
                filter &= " And [Printed By] <> '-'"
            ElseIf vStatusPrint = EnumStatusPrintPO.UnPrintedPOOnly Then
                filter &= " And [Printed By] = '-'"
            End If

            'If vPOno <> "" And filter2 <> "" Then
            '    filter &= " AND (" & filter2 & ")"
            'ElseIf filter = "" Then
            '    filter = filter2
            'End If



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
                    rowDGV.Cells(vFrm.ColAA_PONo.Index).Value = row.No_Transaksi
                    If row.Authorization_Level = 1 Then
                        StatusAuth = "Authorized By Purchase Staff"
                    ElseIf row.Authorization_Level = 2 Then
                        StatusAuth = "Authorized By Purchase Manager"
                    ElseIf row.Authorization_Level = 3 Then
                        StatusAuth = "Authorized By Director"
                    Else
                        StatusAuth = "-"
                    End If
                    rowDGV.Cells(vFrm.ColAA_Status.Index).Value = StatusAuth
                    rowDGV.Cells(vFrm.ColAA_SupplierName.Index).Value = row.Supplier_Name
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

    'aul 
    Public Function IsiDgvBoxBatalOtorisasi(ByVal vLevel As Integer, _
  ByVal Dgv As DataGridView, ByVal TBL As DseTransaksiPO.tbltranspo_authorize_view2_2DataTable, _
  ByVal vDgvUser As DataGridView, _
  Optional ByVal StrictSalesStaff As Boolean = True, Optional ByVal SalesStaff As String = "", _
  Optional ByVal vPOno As String = "", Optional ByVal vDateFrom As Date = #1/1/1950#, _
  Optional ByVal vDateTo As Date = #1/1/1950#, Optional ByVal vPoDate As Boolean = False) As Boolean
        Dim vFrm As New FrmTransaksiPOOtorisasi2
        Dim DTA As New DseTransaksiPOTableAdapters.tbltranspo_authorize_view2_2TableAdapter
        Dim row As DseTransaksiPO.tbltranspo_authorize_view2_2Row
        Dim rowDGV As DataGridViewRow
        Dim filter As String = ""
        Dim filter2 As String = ""
        Dim dv As DataRowView
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None

            'mengisikan semua quotation yang belum diotorisasi
            If vPoDate = False Then
                If StrictSalesStaff = True Then
                    If vLevel = 1 Then
                        DTA.FillBy_UserNAuthorizationLevel_Cancel(TBL, vLevel, SalesStaff)
                    ElseIf vLevel = 2 Then
                        DTA.FillBy_AuthorizationLevel_Cancel_PurchaseManager(TBL, SalesStaff)
                    ElseIf vLevel = 3 Then
                        DTA.FillBy_AuthorizationLevel_Cancel_Director(TBL, SalesStaff)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevel_Cancel(TBL, vLevel)
                End If
            Else
                'filter dengan tanggal
                If StrictSalesStaff = True Then
                    If vLevel = 1 Then
                        DTA.FillBy_UserNAuthorizationLevel_Cancel_TglSampaiDari(TBL, vLevel, SalesStaff, vDateFrom, vDateTo)
                    ElseIf vLevel = 2 Then
                        DTA.FillBy_AuthorizationLevel_Cancel_PurchaseManager_TglSampaiDari(TBL, SalesStaff, vDateFrom, vDateTo)
                    ElseIf vLevel = 3 Then
                        DTA.FillBy_AuthorizationLevel_Cancel_Director_TglSampaiDari(TBL, SalesStaff, vDateFrom, vDateTo)
                    End If
                Else
                    DTA.FillBy_AuthorizationLevel_CancelTglSampaiDari(TBL, vLevel, vDateFrom, vDateTo)
                End If
            End If

            If vPOno <> "" Then
                filter = "[No Transaksi] = '" & vPOno.ToString & "'"
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

            If vPOno <> "" And filter2 <> "" Then
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
                    rowDGV.Cells(vFrm.ColDgvUnAuth_PONo.Index).Value = row.No_Transaksi
                    rowDGV.Cells(vFrm.ColDgvUnAuth_SupplierName.Index).Value = row.Supplier_Name
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

    Public Function GetTotalPurchaseReturnMulti(ByVal vPrimaryPO As Double, ByVal vPrimaryProduct As Double) As Decimal
        Dim dta As New DseTransaksiPOTableAdapters.tbltrans_po4TableAdapter
        Dim NoD As New Nullable(Of Decimal)
        Try
            dta.Connection.ConnectionString = Core_Module.GetConstring
            NoD = dta.ScalarQuery_GetTotalPurchaseReturn(vPrimaryPO, vPrimaryProduct)
            If NoD.HasValue = True Then
                Return NoD.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return 0
        Finally
            dta.Dispose()
        End Try
    End Function
End Class
