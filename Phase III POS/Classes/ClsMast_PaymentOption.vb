Option Explicit On
Option Strict Off
Imports Core_Module
Public Class ClsMast_PaymentOption

    Private mPrimary As Double = "-1"
    Private mCode As String = ""
    Private mDesc As String = ""
    Private mTipePayment As EnumTipePayment = EnumTipePayment.Credit_Card_Payment
    Private mRemarks As String = ""
    Private mBankName As String = ""
    Private mBankAccNumber As String = ""
    Private mMinimumTransAmount As Decimal = 0
    Private mPaymentDelay As Integer = 0
    Private mPaymentJurnal As EnumPaymentJurnal = EnumPaymentJurnal.Automatic
    Private mBankChargeOption As EnumBankChargeOption = EnumBankChargeOption.No_Charge
    Private mChargeRate As EnumChargeRate = EnumChargeRate.Fixed_Nominal_Rate
    Private mCosistentCharge As Integer = 0
    Private mBankChargeNominal As Decimal = 0
    Private mBankChargePercentage As Decimal = 0
    Private mCustChargePercentage As Decimal = 0
    Private mPrimaryAccBankReceivable As Double = "-1"
    Private mPrimaryAccBankCash As Double = "-1"
    Private mPrimaryAccCostPayable As Double = "-1"
    Private mPrimaryAccCreditCardCost As Double = "-1"
    Private mPrimaryAccVoucherCost As Double = "-1"
    Private mMaxVoucherAmount As Decimal = 0
    Private mChkFillApproval As Boolean = False

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianMaster
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Property chkFillApproval() As Boolean
        Get
            Return mChkFillApproval
        End Get
        Set(ByVal value As Boolean)
            mChkFillApproval = value
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

    Public Property Code() As String
        Get
            Return mCode
        End Get
        Set(ByVal value As String)
            mCode = value
        End Set
    End Property

    Public Property Desc() As String
        Get
            Return mDesc
        End Get
        Set(ByVal value As String)
            mDesc = value
        End Set
    End Property

    Public Property TipePayment() As EnumTipePayment
        Get
            Return mTipePayment
        End Get
        Set(ByVal value As EnumTipePayment)
            mTipePayment = value
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

    Public Property BankName() As String
        Get
            Return mBankName
        End Get
        Set(ByVal value As String)
            mBankName = value
        End Set
    End Property

    Public Property BankAccNumber() As String
        Get
            Return mBankAccNumber
        End Get
        Set(ByVal value As String)
            mBankAccNumber = value
        End Set
    End Property

    Public Property MinimumTransAmount() As Decimal
        Get
            Return mMinimumTransAmount
        End Get
        Set(ByVal value As Decimal)
            mMinimumTransAmount = value
        End Set
    End Property

    Public Property PaymentDelay() As Integer
        Get
            Return mPaymentDelay
        End Get
        Set(ByVal value As Integer)
            mPaymentDelay = value
        End Set
    End Property

    Public Property PaymentJurnal() As EnumPaymentJurnal
        Get
            Return mPaymentJurnal
        End Get
        Set(ByVal value As EnumPaymentJurnal)
            mPaymentJurnal = value
        End Set
    End Property

    Public Property BankChargeOption() As EnumBankChargeOption
        Get
            Return mBankChargeOption
        End Get
        Set(ByVal value As EnumBankChargeOption)
            mBankChargeOption = value
        End Set
    End Property

    Public Property ChangeRate() As EnumChargeRate
        Get
            Return mChargeRate
        End Get
        Set(ByVal value As EnumChargeRate)
            mChargeRate = value
        End Set
    End Property

    Public Property ConsistentChange() As Integer
        Get
            Return mCosistentCharge
        End Get
        Set(ByVal value As Integer)
            mCosistentCharge = value
        End Set
    End Property

    Public Property BankChargeNominal() As Decimal
        Get
            Return mBankChargeNominal
        End Get
        Set(ByVal value As Decimal)
            mBankChargeNominal = value
        End Set
    End Property

    Public Property BankChargePercentage() As Decimal
        Get
            Return mBankChargePercentage
        End Get
        Set(ByVal value As Decimal)
            mBankChargePercentage = value
        End Set
    End Property

    Public Property CustChargePercentage() As Decimal
        Get
            Return mCustChargePercentage
        End Get
        Set(ByVal value As Decimal)
            mCustChargePercentage = value
        End Set
    End Property

    Public Property PrimaryAccBankReceivable() As Double
        Get
            Return mPrimaryAccBankReceivable
        End Get
        Set(ByVal value As Double)
            mPrimaryAccBankReceivable = value
        End Set
    End Property

    Public Property PrimaryAccBankCash() As Double
        Get
            Return mPrimaryAccBankCash
        End Get
        Set(ByVal value As Double)
            mPrimaryAccBankCash = value
        End Set
    End Property

    Public Property PrimaryAccCostPayable() As Double
        Get
            Return mPrimaryAccCostPayable
        End Get
        Set(ByVal value As Double)
            mPrimaryAccCostPayable = value
        End Set
    End Property

    Public Property PrimaryAccCreditCardCost() As Double
        Get
            Return mPrimaryAccCreditCardCost
        End Get
        Set(ByVal value As Double)
            mPrimaryAccCreditCardCost = value
        End Set
    End Property

    Public Property PrimaryAccVoucherCost() As Double
        Get
            Return mPrimaryAccVoucherCost
        End Get
        Set(ByVal value As Double)
            mPrimaryAccVoucherCost = value
        End Set
    End Property

    Public Property MaxVoucherAmount() As Decimal
        Get
            Return mMaxVoucherAmount
        End Get
        Set(ByVal value As Decimal)
            mMaxVoucherAmount = value
        End Set
    End Property

    Public Enum EnumTipePayment As Integer
        Credit_Card_Payment
        Debit_Card_Payment
        Cash_Withdrawal
        Voucher_Payment
    End Enum

    Public Enum EnumPaymentJurnal As Integer
        Automatic
        Manual
    End Enum

    Public Enum EnumBankChargeOption As Integer
        No_Charge
        Apply_Charge_For_Customer
        Apply_Charge_As_Internal_Cost
    End Enum

    Public Enum EnumChargeRate As Integer
        Fixed_Nominal_Rate
        Percentage
    End Enum

    Public Property CreatedBy() As String
        Get
            Return Me.mCreatedBy
        End Get
        Set(ByVal value As String)
            mCreatedBy = value
        End Set
    End Property

    Public Property CreatedDate() As Date
        Get
            Return Me.mCreatedDate
        End Get
        Set(ByVal value As Date)
            mCreatedDate = value
        End Set
    End Property

    Public Property ModifiedBy() As String
        Get
            Return Me.mModifiedBy
        End Get
        Set(ByVal value As String)
            mModifiedBy = value
        End Set
    End Property

    Public Property ModifiedDate() As Date
        Get
            Return Me.mModifiedDate
        End Get
        Set(ByVal value As Date)
            mModifiedDate = value
        End Set
    End Property

    Public Property TimeStamp() As Date
        Get
            Return Me.mTimestamp
        End Get
        Set(ByVal value As Date)
            Me.mTimestamp = value
        End Set
    End Property

    Public Function GetConstring() As String
        Try
            Return "server=" + My.Settings.server + ";user id=" + My.Settings.user_id + ";Password=" + My.Settings.password + _
            ";persist security info=True;database=" + My.Settings.database_name + ";port=" + My.Settings.port_number + ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "server=localhost;user id=root;Password=123;persist security info=True;database=core_db;port=3307"
        End Try
    End Function

    Public Sub New()
        Try
            Me.DTAAntrian.Connection.ConnectionString = Me.GetConstring
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub GetNoAntrian()
        Try
            DTAAntrian.Connection.ConnectionString = Me.GetConstring
            Me.vIdentifier = Me.ObjRnd.Next(1000000)
            'jangan sampai sama dengan identifier yang sudah ada
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
            End While
            'masukan data baru ke antrian
            Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)

            'cek jangan sampai ada 2 data yang sama
            While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 1
                'bila ada yang sama maka harus masukkan data baru
                DTAAntrian.DeleteQuery(Me.vTipeAntrian, Me.vIdentifier)
                'minta identifier baru
                Me.vIdentifier = Me.ObjRnd.Next(1000000)
                'jangan sampai sama dengan identifier yang sudah ada
                While Me.DTAAntrian.ScalarQuery_GetJmlIdentifier(Me.vTipeAntrian, Me.vIdentifier).Value > 0
                    Me.vIdentifier = Me.ObjRnd.Next(1000000)
                    'masukan data baru ke antrian
                    Me.DTAAntrian.Insert(Me.vTipeAntrian, Me.vIdentifier)
                End While
            End While

            'ambil no antrian
            Me.vPrimary_Antrian = DTAAntrian.ScalarQuery_GetPrimary(Me.vIdentifier, Me.vTipeAntrian).Value
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetPrimary(ByVal vCode As String) As Double
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCode(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCode(vPrimary)
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

    Public Function GetDesc(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary)
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

    Public Function GetCreatedBy(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCreatedBy(vPrimary)
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

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary).Value
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTipePayment(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetTipePayment(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTipePayment(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlByCode(ByVal vCode As String) As Double
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCode(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vCode As String) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCode(vCode).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCode(vCode).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vObjCD As ClsCekData, ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim ListPrimaryAkun As New List(Of Double)
        Try
            'tidak boleh ada kode yang sama
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mCode) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Duplication in code, please change your code", _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Duplikasi code, silahkan ubah code", _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            Else
                Dim CodeLama As String = Me.GetCode(Me.mPrimary)
                If CodeLama <> Me.mCode Then
                    If Me.ApakahSudahAda(Me.mCode) = True Then
                        If vShowMsg = True Then
                            If vEnumBhs = EnumBahasa.Inggris Then
                                MsgBox("Duplication in code, please change your code", _
                                MsgBoxStyle.Information, "Information")
                            ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                                MsgBox("Duplikasi code, silahkan ubah code", _
                                MsgBoxStyle.Information, "Informasi")
                            End If
                        End If
                        Return False
                    End If
                End If
            End If

            'yang harus diisikan:
            '1. kode
            If Me.mCode = "" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input code.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan code.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If
            '2. nama
            If Me.mDesc = "" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input description.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan decription.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If


            ' Cek Akun yang dipergunakan 
            If Me.mTipePayment = EnumTipePayment.Credit_Card_Payment Or _
            Me.mTipePayment = EnumTipePayment.Debit_Card_Payment Then
                If ListPrimaryAkun.Contains(Me.mPrimaryAccBankReceivable) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Bank receivable account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Bank receivable account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccBankReceivable)
                End If
                If ListPrimaryAkun.Contains(Me.mPrimaryAccBankCash) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Bank cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Bank cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccBankCash)
                End If
                If ListPrimaryAkun.Contains(Me.mPrimaryAccCostPayable) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Cost payable code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Cost payable code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccCostPayable)
                End If
                If ListPrimaryAkun.Contains(Me.mPrimaryAccCreditCardCost) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Credit card cost code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Credit card cost code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccCreditCardCost)
                End If
            End If
            ''
            If Me.mTipePayment = EnumTipePayment.Cash_Withdrawal Then
                If ListPrimaryAkun.Contains(Me.mPrimaryAccBankReceivable) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Bank receivable account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Bank receivable account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccBankReceivable)
                End If
                If ListPrimaryAkun.Contains(Me.mPrimaryAccBankCash) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Bank cash account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Bank cash account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccBankCash)
                End If
            End If
            ''
            If Me.mTipePayment = EnumTipePayment.Voucher_Payment Then
                If ListPrimaryAkun.Contains(Me.mPrimaryAccVoucherCost) = True Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Cost account code has been used in another account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Cost account code sudah digunakan di account code lain.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                Else
                    ListPrimaryAkun.Add(Me.mPrimaryAccVoucherCost)
                End If
            End If

            '4 Remarks 
            If String.IsNullOrEmpty(Me.mRemarks) = False Then
                If vObjCD.CekDataKosong(Me.mRemarks) = False Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Data can not contain the following characters: ' , % , # , or * .", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data tidak boleh memuat karakter: ' , % , #, atau * .", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ListPrimaryAkun = Nothing
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictEdit As Boolean, _
    ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictEdit = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If

                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can edit this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat diedit oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, ByVal vStrictDelete As Boolean, _
    ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Dim pCreatedBy As String = ""
        Dim ObjPOS As New ClsTrans_POS
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            '1. tidak boleh hapus jika melanggar aturan strict edit
            If vStrictDelete = True Then
                If IsNothing(DTA.ScalarQuery_GetCreatedBy(vPrimary)) = False Then
                    pCreatedBy = DTA.ScalarQuery_GetCreatedBy(vPrimary)
                Else
                    pCreatedBy = ""
                End If
                If pCreatedBy <> vUserCode Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("User can delete this data: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Data ini hanya dapat didelete oleh: " + pCreatedBy, _
                            MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    Return False
                End If
            End If

            '2. Trans POS
            If ObjPOS.GetJmlByPrimaryPayment(vPrimary) > 0 Then
                MsgBox("Data cannot be deleted because there are other transaction that refering this data", MsgBoxStyle.Information, "Information")
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            ObjPOS = Nothing
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMast_Payment_Option.tblmaster_payment_option1DataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel, _
    Optional ByVal vBwtGeneralSetting As Boolean = False) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option1TableAdapter
        Dim Row As DseMast_Payment_Option.tblmaster_payment_option1Row
        Dim QueryResult As Double
        Dim StartT As Date
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None

            If SearchCriteria = EnumSearchCriteria.ByCode Then
                QueryResult = DTA.FillBy_CodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByRemarks Then
                QueryResult = DTA.FillBy_RemarksLIKE(TBL, Criteria + "%")
            End If

            If vBwtGeneralSetting = True Then
                Row = TBL.NewRow
                With Row
                    .primary = -1
                    .code = "CASH"
                    .desc = "Cash Payment"
                    .remarks = ""
                End With
                TBL.Addtblmaster_payment_option1Row(Row)
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(1).HeaderText = "Code"
                DGV.Columns(2).HeaderText = "Description"
                DGV.Columns(3).HeaderText = "Remarks"
            End If


            Dim TS As TimeSpan = Now - StartT
            If QueryResult > 0 Then
                If TS.Ticks > TimeSpan.TicksPerSecond Then
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(TS.Seconds) + " seconds"
                Else
                    tsl.Text = CStr(QueryResult) + " result(s) in " + CStr(Math.Round(TS.Ticks / TimeSpan.TicksPerMillisecond, _
                    0, MidpointRounding.AwayFromZero)) + " miliseconds"
                End If
            Else
                tsl.Text = "No result found"
            End If

            DGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            DGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option2TableAdapter
        Dim TBL As New DseMast_Payment_Option.tblmaster_payment_option2DataTable
        Dim row As DseMast_Payment_Option.tblmaster_payment_option2Row
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .primary
                    Me.mCode = .code
                    Me.mDesc = .desc
                    Me.mTipePayment = .tipe_payment
                    Me.mRemarks = .remarks
                    Me.mBankName = .bank_name
                    Me.mBankAccNumber = .bank_acc_number
                    Me.mMinimumTransAmount = .minimum_transaction_amount
                    Me.mPaymentDelay = .payment_delay
                    Me.mPaymentJurnal = .payment_jurnal
                    Me.mBankChargeOption = .bank_charge_option
                    Me.mChargeRate = .charge_rate
                    Me.mCosistentCharge = .consistent_charge
                    Me.mBankChargeNominal = .bank_charge_nominal
                    Me.mBankChargePercentage = .bank_charge_percentage
                    Me.mCustChargePercentage = .cust_charge_percentage
                    Me.mPrimaryAccBankReceivable = .primary_account_bank_receivable
                    Me.mPrimaryAccBankCash = .primary_account_bank_cash
                    Me.mPrimaryAccCostPayable = .primary_account_cost_payable
                    Me.mPrimaryAccCreditCardCost = .primary_account_cc_cost
                    Me.mPrimaryAccVoucherCost = .primary_account_voucher_cost
                    Me.mMaxVoucherAmount = .maximum_voucher_amount
                    Me.mChkFillApproval = .Fill_Approval_Code

                    Me.mCreatedBy = .created_by
                    Me.mCreatedDate = .created_date
                    Me.mModifiedBy = .modified_by
                    Me.mModifiedDate = .modified_date
                    Me.mTimestamp = .timestamp
                End With
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
        End Try
    End Function

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As EnumSearchCriteria) As Double
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option1TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If vSearchBy = EnumSearchCriteria.ByCode Then
                Return DTA.ScalarQuery_GetJmlByCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByRemarks Then
                Return DTA.ScalarQuery_GetJmlByRemraksLIKE(vKriteria + "%")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmMast_PaymentOption) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanBaru1(vUserCode, vPrimary, vObjCD, _
                        vEnumBhs, i)

                        'hapus data antri untuk transaksi ini
                        If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanBaru1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Dim DTA2 As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option2TableAdapter
        Dim vApproval As Integer = 0
        Try
            DTA.Connection.ConnectionString = Me.GetConstring

            If Me.ApakahDataBenar(True, True, vObjCD, vEnumBhs, i) = True Then
                If Me.mChkFillApproval = True Then
                    vApproval = 1
                Else
                    vApproval = 0
                End If
                If DTA2.Insert(Me.mCode, Me.mDesc, Me.mTipePayment, Me.mRemarks, _
                Me.mBankName, Me.mBankAccNumber, Me.mMinimumTransAmount, Me.mPaymentDelay, Me.mPaymentJurnal, Me.mBankChargeOption, _
                Me.mChargeRate, Me.mCosistentCharge, Me.mBankChargeNominal, Me.mBankChargePercentage, Me.mCustChargePercentage, _
                Me.mPrimaryAccBankReceivable, Me.mPrimaryAccBankCash, Me.mPrimaryAccCostPayable, Me.mPrimaryAccCreditCardCost, _
                Me.mPrimaryAccVoucherCost, Me.mMaxVoucherAmount, _
                vUserCode, Now, vUserCode, Now, Now, vApproval) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCode)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode

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
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.SimpanPerubahan1(vUserCode, vPrimary, vObjCD, vEnumBhs, i)

                        'hapus data antri untuk transaksi ini
                        If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function SimpanPerubahan1(ByVal vUserCode As String, _
    ByVal vPrimary As Double, ByVal vObjCD As ClsCekData, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Dim DTA2 As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option2TableAdapter
        Dim vApproval As Integer = 0
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(False, True, vObjCD, vEnumBhs, i) = True Then
                If Me.mChkFillApproval = True Then
                    vApproval = 1
                Else
                    vApproval = 0
                End If
                If DTA2.Update(Me.mCode, Me.mDesc, Me.mTipePayment, Me.mRemarks, _
                Me.mBankName, Me.mBankAccNumber, Me.mMinimumTransAmount, Me.mPaymentDelay, Me.mPaymentJurnal, Me.mBankChargeOption, _
                Me.mChargeRate, Me.mCosistentCharge, Me.mBankChargeNominal, Me.mBankChargePercentage, Me.mCustChargePercentage, _
                Me.mPrimaryAccBankReceivable, Me.mPrimaryAccBankCash, Me.mPrimaryAccCostPayable, Me.mPrimaryAccCreditCardCost, _
                Me.mPrimaryAccVoucherCost, Me.mMaxVoucherAmount, _
                vUserCode, Now, Now, vApproval, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA2.Dispose()
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim CheckInterval As Double = 0
        Dim TryCounter As Integer = 0
        Dim Start As Double = 0
        Dim retval As Boolean = True
        Try
            'ambil no antrian
            Me.GetNoAntrian()

            'coba simpan data baru
            Me.Clock_Antri.Reset()
            Me.Clock_Antri.Start()
            While Start = 0
                'hanya mencoba ke database setelah interval waktu tertentu saja
                If Me.Clock_Antri.ElapsedMilliseconds > CheckInterval Then
                    If Me.DTAAntrian.ScalarQuery_GetPrimaryMin(Me.vTipeAntrian).Value = Me.vPrimary_Antrian Then
                        'clock antri matikan
                        Me.Clock_Antri.Stop()
                        'sudah saatnya melakukan proses simpan baru
                        retval = Me.Hapus1(vPrimary, vUserCode, _
                        vStrictDelete, vShowMsg, vEnumBhs, i)
                        'hapus data antri untuk transaksi ini
                        If Me.DTAAntrian.Delete(Me.vPrimary_Antrian) = 0 Then
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
            End While
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function Hapus1(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmMast_PaymentOption) As Boolean
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, _
                vStrictDelete, vShowMsg, vEnumBhs, i) = True Then
                'Hapus
                DTA.Delete(vPrimary)
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

    Public Sub ResetObject()
        Try
            mPrimary = -1
            mCode = ""
            mDesc = ""
            mTipePayment = EnumTipePayment.Credit_Card_Payment
            mRemarks = ""
            mBankName = ""
            mBankAccNumber = ""
            mMinimumTransAmount = 0
            mPaymentDelay = 0
            mPaymentJurnal = EnumPaymentJurnal.Automatic
            mBankChargeOption = EnumBankChargeOption.No_Charge
            mChargeRate = EnumChargeRate.Fixed_Nominal_Rate
            mCosistentCharge = 0
            mBankChargeNominal = 0
            mBankChargePercentage = 0
            mCustChargePercentage = 0
            mPrimaryAccBankReceivable = "-1"
            mPrimaryAccBankCash = "-1"
            mPrimaryAccCostPayable = "-1"
            mPrimaryAccCreditCardCost = "-1"
            mPrimaryAccVoucherCost = "-1"
            mMaxVoucherAmount = 0
            mChkFillApproval = False

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetUserPrecentage(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetUserPrecentage(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetUserPrecentage(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiCBOPaymentOption(ByVal vCbo As ComboBox, ByVal vTBL As DseMast_Payment_Option.tblmaster_payment_option1_3DataTable, _
    ByVal DTA As DseMast_Payment_OptionTableAdapters.tblmaster_payment_option1_3TableAdapter, ByVal vPrimaryPayment As Double, _
    Optional ByVal vIsSinglePayment As Boolean = False)
        Dim RowPayment As DseMast_Payment_Option.tblmaster_payment_option1_3Row
        Dim RowView As Data.DataRowView
        Dim Index As Integer = 0
        Try
            'get connection
            DTA.Connection.ConnectionString = Me.GetConstring
            'kosongkan cbo
            vCbo.DataSource = Nothing
            vCbo.Items.Clear()

            'isi tabel
            If vIsSinglePayment = True Then
                DTA.FillBy_TidakWithdrawal(vTBL)
            Else
                DTA.Fill(vTBL)
            End If

            'Tambahkan Type CASH Payment, Primary = -1
            RowPayment = vTBL.NewRow
            With RowPayment
                .primary = -1
                .code = "Cash"
            End With
            vTBL.Addtblmaster_payment_option1_3Row(RowPayment)

            'hubungkan dengan cbo
            vCbo.DataSource = vTBL.DefaultView
            vTBL.DefaultView.Sort = "[code] ASC"
            vCbo.DisplayMember = "code"

            'Pilih Code
            For Each RowView In vTBL.DefaultView
                RowPayment = RowView.Row
                If RowPayment.primary = vPrimaryPayment Then
                    Exit For
                Else
                    Index += 1
                End If
            Next
            vCbo.SelectedIndex = Index
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetFillApprovalCode(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_option2TableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetFillApprovalCode(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetFillApprovalCode(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetTypePayment(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPaymentType(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetPaymentType(vPrimary).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBankChargeNominal(ByVal vPrimary As Double) As Decimal
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetBankChargeNominal(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetBankChargeNominal(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBankAccountNumber(ByVal vPrimary As Double) As String
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetBankAcc(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetBankAcc(vPrimary)
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

    Public Function GetBankChargeOption(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMast_Payment_OptionTableAdapters.tblmaster_payment_optionTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetBankChargeOption(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetBankChargeOption(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
