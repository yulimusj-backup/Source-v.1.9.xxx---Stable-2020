Option Explicit On
Option Strict Off
Imports Core_Module

Public Class ClsTrans_POS_GeneralSetting

    Private mPrimary As Double = -1
    Private mCompanyName As String = ""
    Private mShowCompanyName As Integer = 0
    Private mShowCompanyLogo As Integer = 0
    Private mCompanyLogo As String = ""
    Private mPrimaryCurr As Double = -1
    Private mPrimaryPayment As Double = -1
    Private mAllowCreditCardPayment As Integer = 0
    Private mAllowDebitCardPayment As Integer = 0
    Private mAllowCashPayment As Integer = 0
    Private mAllowVoucherPayment As Integer = 0
    Private mUseTaxRate As Integer = 0
    Private mPersenTaxRate As Decimal = 0
    Private mAllowRounding As Integer = 0
    Private mRoundingTipe As EnumRoundingTipe = EnumRoundingTipe.Up
    Private mRoundingValue As EnumRoundingValue = EnumRoundingValue.G_Satu
    Private mRoundingOn As EnumRoundingOn = EnumRoundingOn.Before_Tax_Calculation
    Private mPrimaryAcc As Double = -1
    Private mAutomaticTransfer As EnumAutomaticTransfer = EnumAutomaticTransfer.Stop_Transfer
    Private mExpiryASC As EnumExpiryASC = EnumExpiryASC.Descending
    Private mBatchASC As EnumBatchASC = EnumBatchASC.Descending
    Private mAllowLocationControl As Integer = 0
    Private mAllowBatchNumberControl As Integer = 0
    Private mAllowExpiryControl As Integer = 0
    Private mInputLocationOnCashierForm As EnumInputLocationOnCashierForm = EnumInputLocationOnCashierForm.No
    Private mDefPaymentDiskon As String = ""
    Private mUsePaymentDiskon As Integer = 0

    Dim mCreatedBy As String = ""
    Dim mCreatedDate As Date = Now
    Dim mModifiedBy As String = ""
    Dim mModifiedDate As Date = Now
    Dim mTimestamp As Date = Now

    'elemen stock yang digunakan untuk antrian simpan data
    Dim vPrimary_Antrian As Double
    Dim vIdentifier As Double
    Dim vTipeAntrian As Core_Module.ClsAntrian.EnumTipeAntrian = Core_Module.ClsAntrian.EnumTipeAntrian.AntrianTransaksiPOS
    Dim ObjRnd As New Random
    Dim WithEvents Clock_Antri As New Stopwatch
    Dim DTAAntrian As New DseTransaksiAntrianTableAdapters.tbltrans_antrianTableAdapter

    Public Property UsePaymentDiskon() As Integer
        Get
            Return mUsePaymentDiskon
        End Get
        Set(ByVal value As Integer)
            mUsePaymentDiskon = value
        End Set
    End Property

    Public Property DefPaymentDiskon() As String
        Get
            Return mDefPaymentDiskon
        End Get
        Set(ByVal value As String)
            mDefPaymentDiskon = value
        End Set
    End Property

    Public Property AllowLocationControl() As Integer
        Get
            Return mAllowLocationControl
        End Get
        Set(ByVal value As Integer)
            mAllowLocationControl = value
        End Set
    End Property

    Public Property AllowBatchNumberControl() As Integer
        Get
            Return mAllowBatchNumberControl
        End Get
        Set(ByVal value As Integer)
            mAllowBatchNumberControl = value
        End Set
    End Property

    Public Property AllowExpiryControl() As Integer
        Get
            Return mAllowExpiryControl
        End Get
        Set(ByVal value As Integer)
            mAllowExpiryControl = value
        End Set
    End Property

    Public Enum EnumInputLocationOnCashierForm As Integer
        No
        Yes
    End Enum

    Public Property InputLocationOnCashierForm() As EnumInputLocationOnCashierForm
        Get
            Return mInputLocationOnCashierForm
        End Get
        Set(ByVal value As EnumInputLocationOnCashierForm)
            mInputLocationOnCashierForm = value
        End Set
    End Property

    Public Enum EnumBatchASC As Integer
        Acscending
        Descending
    End Enum

    Public Property BatchASC() As EnumExpiryASC
        Get
            Return mBatchASC
        End Get
        Set(ByVal value As EnumExpiryASC)
            mBatchASC = value
        End Set
    End Property

    Public Enum EnumExpiryASC As Integer
        Acscending
        Descending
    End Enum

    Public Property ExpiryASC() As EnumExpiryASC
        Get
            Return mExpiryASC
        End Get
        Set(ByVal value As EnumExpiryASC)
            mExpiryASC = value
        End Set
    End Property

    Public Enum EnumAutomaticTransfer As Integer
        Stop_Transfer
        Automatically
    End Enum

    Public Property AutomaticTransfer() As EnumAutomaticTransfer
        Get
            Return mAutomaticTransfer
        End Get
        Set(ByVal value As EnumAutomaticTransfer)
            mAutomaticTransfer = value
        End Set
    End Property

    Public Property PrimaryPayment() As Double
        Get
            Return mPrimaryPayment
        End Get
        Set(ByVal value As Double)
            mPrimaryPayment = value
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

    Public Property CompanyName() As String
        Get
            Return mCompanyName
        End Get
        Set(ByVal value As String)
            mCompanyName = value
        End Set
    End Property

    Public Property ShowCompanyName() As Integer
        Get
            Return mShowCompanyName
        End Get
        Set(ByVal value As Integer)
            mShowCompanyName = value
        End Set
    End Property

    Public Property ShowCompanyLogo() As Integer
        Get
            Return mShowCompanyLogo
        End Get
        Set(ByVal value As Integer)
            mShowCompanyLogo = value
        End Set
    End Property

    Public Property CompanyLogo() As String
        Get
            Return mCompanyLogo
        End Get
        Set(ByVal value As String)
            mCompanyLogo = value
        End Set
    End Property

    Public Property PrimaryCurr() As Double
        Get
            Return mPrimaryCurr
        End Get
        Set(ByVal value As Double)
            mPrimaryCurr = value
        End Set
    End Property

    Public Property AllowCreditCardPayment() As Integer
        Get
            Return mAllowCreditCardPayment
        End Get
        Set(ByVal value As Integer)
            mAllowCreditCardPayment = value
        End Set
    End Property

    Public Property AllowDebitCardPayment() As Integer
        Get
            Return mAllowDebitCardPayment
        End Get
        Set(ByVal value As Integer)
            mAllowDebitCardPayment = value
        End Set
    End Property

    Public Property AllowCashPayment() As Integer
        Get
            Return mAllowCashPayment
        End Get
        Set(ByVal value As Integer)
            mAllowCashPayment = value
        End Set
    End Property

    Public Property AllowVoucherPayment() As Integer
        Get
            Return mAllowVoucherPayment
        End Get
        Set(ByVal value As Integer)
            mAllowVoucherPayment = value
        End Set
    End Property

    Public Property UseTaxRate() As Integer
        Get
            Return mUseTaxRate
        End Get
        Set(ByVal value As Integer)
            mUseTaxRate = value
        End Set
    End Property

    Public Property PersenTaxRate() As Decimal
        Get
            Return mPersenTaxRate
        End Get
        Set(ByVal value As Decimal)
            mPersenTaxRate = value
        End Set
    End Property

    Public Property AllowRounding() As Integer
        Get
            Return mAllowRounding
        End Get
        Set(ByVal value As Integer)
            mAllowRounding = value
        End Set
    End Property

    Public Property PrimaryAcc() As Double
        Get
            Return mPrimaryAcc
        End Get
        Set(ByVal value As Double)
            mPrimaryAcc = value
        End Set
    End Property

    Public Enum EnumRoundingOn As Integer
        Before_Tax_Calculation
        After_Tax_Calculation
    End Enum

    Public Property RoundingOn() As EnumRoundingOn
        Get
            Return mRoundingOn
        End Get
        Set(ByVal value As EnumRoundingOn)
            mRoundingOn = value
        End Set
    End Property

    Public Enum EnumRoundingValue As Integer
        A_Seribu
        B_LimaRatus
        C_Seratus
        D_LimaPuluh
        E_Sepuluh
        F_Lima
        G_Satu
        H_NolKomaLima
        I_NolKomaSatu
        J_NolKomaNolLima
        K_NolKomaNolSatu
        L_NolKomaNolNolLima
        M_NolKomaNolNolSatu
    End Enum

    Public Property RoundingValue() As EnumRoundingValue
        Get
            Return mRoundingValue
        End Get
        Set(ByVal value As EnumRoundingValue)
            mRoundingValue = value
        End Set
    End Property

    Public Enum EnumRoundingTipe As Integer
        Up
        Down
        Middle
    End Enum

    Public Property RoundingTipe() As EnumRoundingTipe
        Get
            Return mRoundingTipe
        End Get
        Set(ByVal value As EnumRoundingTipe)
            mRoundingTipe = value
        End Set
    End Property

    Private mAllowToCancel As Integer = 0
    Private mUseAuthorToCancel As Integer = 0

    Public Property AllowToCancel() As Integer
        Get
            Return mAllowToCancel
        End Get
        Set(ByVal value As Integer)
            mAllowToCancel = value
        End Set
    End Property

    Public Property UseAuthorToCancel() As Integer
        Get
            Return mUseAuthorToCancel
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToCancel = value
        End Set
    End Property

    Private mAllowUserToCancelWholeTrans As Integer = 0
    Private mUseAuthorToCancelWholeTrans As Integer = 0

    Public Property AllowUserToCancelWholeTrans() As Integer
        Get
            Return mAllowUserToCancelWholeTrans
        End Get
        Set(ByVal value As Integer)
            mAllowUserToCancelWholeTrans = value
        End Set
    End Property

    Public Property UseAuthorToCancelWholeTrans() As Integer
        Get
            Return mUseAuthorToCancelWholeTrans
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToCancelWholeTrans = value
        End Set
    End Property

    Private mAllowUserToModifyTrans As Integer = 0
    Private mUseAuthorToModifyTrans As Integer = 0

    Public Property AllowUserToModifyTrans() As Integer
        Get
            Return mAllowUserToModifyTrans
        End Get
        Set(ByVal value As Integer)
            mAllowUserToModifyTrans = value
        End Set
    End Property

    Public Property UseAuthorToModifyTrans() As Integer
        Get
            Return mUseAuthorToModifyTrans
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToModifyTrans = value
        End Set
    End Property

    Private mAllowUserToModifyUnitPrice As Integer = 0
    Private mUseAuthorToModifyUnitPrice As Integer = 0

    Public Property AllowUserToModifyUnitPrice() As Integer
        Get
            Return mAllowUserToModifyUnitPrice
        End Get
        Set(ByVal value As Integer)
            mAllowUserToModifyUnitPrice = value
        End Set
    End Property

    Public Property UseAuthorToModifyUnitPrice() As Integer
        Get
            Return mUseAuthorToModifyUnitPrice
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToModifyUnitPrice = value
        End Set
    End Property

    Private mAllowUserToModifyDiscountRate As Integer = 0
    Private mUseAuthorToModifyDiscountRate As Integer = 0

    Public Property AllowUserToModifyDiscountRate() As Integer
        Get
            Return mAllowUserToModifyDiscountRate
        End Get
        Set(ByVal value As Integer)
            mAllowUserToModifyDiscountRate = value
        End Set
    End Property

    Public Property UseAuthorToModifyDiscountRate() As Integer
        Get
            Return mUseAuthorToModifyDiscountRate
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToModifyDiscountRate = value
        End Set
    End Property

    Private mAllowUserToReprintReceipt As Integer = 0
    Private mUseAuthorToReprintReceipt As Integer = 0

    Public Property AllowUserToReprintReceipt() As Integer
        Get
            Return mAllowUserToReprintReceipt
        End Get
        Set(ByVal value As Integer)
            mAllowUserToReprintReceipt = value
        End Set
    End Property

    Public Property UseAuthorToReprintReceipt() As Integer
        Get
            Return mUseAuthorToReprintReceipt
        End Get
        Set(ByVal value As Integer)
            mUseAuthorToReprintReceipt = value
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

    ''Coding standarnya 
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

    Public Function GetJmlByCompanyName(ByVal vCompanyName As String) As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCompanyName(vCompanyName).HasValue Then
                Return DTA.ScalarQuery_GetJmlByCompanyName(vCompanyName).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimary(ByVal vCompanyName As String) As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimary(vCompanyName).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vCompanyName).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetCompanyName(ByVal vPrimary As Double) As String
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetCompanyName(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetCompanyName(vPrimary)
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
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
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
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
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

    Public Function ApakahSudahAda(ByVal vCompanyName As String) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJmlByCompanyName(vCompanyName).Value > 0 Then
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

    ''' <summary>
    ''' data ini digunakan untuk mengecek kebenaran data yang diberikan
    ''' </summary>
    ''' <param name="vBaru"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ApakahDataBenar(ByVal vBaru As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vDGVItem As DataGridView, ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
        Try
            'yang harus diisikan:
            '1. kode
            If Me.mCompanyName = "" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input company name.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan company name.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '2. Default Currency
            If Me.mPrimaryCurr = "-1" Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input default currency.", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan default currency.", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                Return False
            End If

            '3 Payment Option
            'If Me.mPrimaryPayment = "-1" Then
            'If vShowMsg = True Then
            'If vEnumBhs = EnumBahasa.Inggris Then
            'MsgBox("Please input payment option.", MsgBoxStyle.Information, "Information")
            'ElseIf vEnumBhs = EnumBahasa.Indonesia Then
            'MsgBox("Masukan payment option.", MsgBoxStyle.Information, "Informasi")
            'End If
            'End If
            'Return False
            'End If

            '4. Allow Rounding 
            If i.chkAllowRounding.Checked = True Then
                If i.TxtRounding_PrimaryAccount.Text = "-1" Then
                    If vShowMsg = True Then
                        If vEnumBhs = EnumBahasa.Inggris Then
                            MsgBox("Please input account code.", MsgBoxStyle.Information, "Information")
                        ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                            MsgBox("Masukan account code.", MsgBoxStyle.Information, "Informasi")
                        End If
                    End If
                    i.TabControl2.SelectedIndex = 2
                    Return False
                End If
            End If

            'Item tidak boleh kosong 
            i.dgvAutorization.AllowUserToAddRows = False
            If i.dgvAutorization.Rows.Count <= 0 Then
                If vShowMsg = True Then
                    If vEnumBhs = EnumBahasa.Inggris Then
                        MsgBox("Please input item data", MsgBoxStyle.Information, "Information")
                    ElseIf vEnumBhs = EnumBahasa.Indonesia Then
                        MsgBox("Masukan data item", MsgBoxStyle.Information, "Informasi")
                    End If
                End If
                i.TabControl2.SelectedIndex = 1
                Return False
            End If
            i.dgvAutorization.AllowUserToAddRows = True

            'kalo mpe sini ok
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahBolehEdit(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictEdit As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Return True
            Exit Try

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Dim pCreatedBy As String = ""
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            Return True
            Exit Try

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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function CariDetail(ByVal vPrimary As Double, ByVal vDGVItem As DataGridView, _
    ByVal vObjCD As ClsCekData, ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_SAVETableAdapter
        Dim TBL As New DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_SAVEDataTable
        Dim row As DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_SAVERow
        Dim ObjItem As New ClsTrans_POS_GeneralSetting_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            DTA.FillBy_Primary(TBL, vPrimary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mCompanyName = .Company_Name
                    Me.mCompanyLogo = .Company_Logo
                    Me.mShowCompanyName = .Show_Company_Name
                    Me.mShowCompanyLogo = .Show_Company_Logo
                    Me.mPrimaryCurr = .Primary_Currency
                    Me.mAllowCreditCardPayment = .Allow_Credit_Card_Payment
                    Me.mAllowDebitCardPayment = .Allow_Debit_Card_Payment
                    Me.mAllowCashPayment = .Allow_Cash_Withdrawal
                    Me.mAllowVoucherPayment = .Allow_Voucher_Payment
                    Me.mUseTaxRate = .Use_Tax_Rate
                    Me.mPersenTaxRate = .Persen_Tax_Rate
                    Me.mAllowRounding = .Allow_Rounding
                    Me.mRoundingTipe = .Rounding_Tipe
                    Me.mRoundingValue = .Rounding_Value
                    Me.mRoundingOn = .Rounding_On
                    Me.mPrimaryAcc = .Primary_Account
                    Me.mAllowToCancel = .Allow_User_To_Cancel
                    Me.mUseAuthorToCancel = .Use_Authorization_To_Cancel
                    Me.mAllowUserToCancelWholeTrans = .Allow_User_To_Cancel_Whole_Transaction
                    Me.mUseAuthorToCancelWholeTrans = .Use_Authorization_To_Cancel_Whole_Transaction
                    Me.mAllowUserToModifyTrans = .Allow_User_To_Modify_Transaction
                    Me.mUseAuthorToModifyTrans = .Use_Authorization_To_Modify_Transaction
                    Me.mAllowUserToModifyUnitPrice = .Allow_User_To_Modify_Unit_Price
                    Me.mUseAuthorToModifyUnitPrice = .Use_Authorization_To_Modify_Unit_Price
                    Me.mAllowUserToModifyDiscountRate = .Allow_User_To_Modify_Discount_Rate
                    Me.mUseAuthorToModifyDiscountRate = .Use_Authorization_To_Modify_Discount_Rate
                    Me.mAllowUserToReprintReceipt = .Allow_User_To_Reprint_Receipt
                    Me.mUseAuthorToReprintReceipt = .Use_Authorization_To_Reprint_Receipt
                    Me.mPrimaryPayment = .Primary_Payment
                    Me.mAutomaticTransfer = .Automatically_Resolve_Transfer_Problems
                    Me.mExpiryASC = .Expiry_ASC
                    Me.mBatchASC = .Batch_ASC
                    Me.mAllowLocationControl = .Allow_Location_Control
                    Me.mAllowBatchNumberControl = .Allow_Batch_Number_Control
                    Me.mAllowExpiryControl = .Allow_Expiry_Control
                    Me.mInputLocationOnCashierForm = .Input_Location_On_Cashier_Form
                    Me.mDefPaymentDiskon = .Default_Payment_Disc
                    Me.mUsePaymentDiskon = .Modify_Payment_Disc

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
                End With

                'cari data item
                ObjItem.IsiDGV(Primary, vDGVItem, i, vObjCD)
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanBaru(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, ByVal vEnumBhs As EnumBahasa, _
    ByVal i As FrmTrans_POS_Settings) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Dim DTA_SAVE As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_SAVETableAdapter
        Dim ObjItem As New ClsTrans_POS_GeneralSetting_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(True, True, i.dgvAutorization, vEnumBhs, i) = True Then
                If DTA_SAVE.Insert(Me.mCompanyName, Me.mShowCompanyName, Me.mShowCompanyLogo, Me.mCompanyLogo, _
                Me.mPrimaryCurr, Me.mAllowCreditCardPayment, Me.mAllowDebitCardPayment, Me.mAllowCashPayment, _
                Me.mAllowVoucherPayment, Me.mUseTaxRate, Me.mPersenTaxRate, Me.mAllowRounding, Me.mRoundingTipe, _
                Me.mRoundingValue, Me.mRoundingOn, Me.mPrimaryAcc, Me.mAllowToCancel, Me.mUseAuthorToCancel, _
                Me.mAllowUserToCancelWholeTrans, Me.mUseAuthorToCancelWholeTrans, Me.mAllowUserToModifyTrans, _
                Me.mUseAuthorToModifyTrans, Me.mAllowUserToModifyUnitPrice, Me.mUseAuthorToModifyUnitPrice, _
                Me.mAllowUserToModifyDiscountRate, Me.mUseAuthorToModifyDiscountRate, Me.mAllowUserToReprintReceipt, _
                Me.mUseAuthorToReprintReceipt, vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryPayment, Me.mAutomaticTransfer, _
                Me.mExpiryASC, Me.mBatchASC, Me.mAllowLocationControl, Me.mAllowBatchNumberControl, Me.mAllowExpiryControl, _
                Me.mInputLocationOnCashierForm, Me.mDefPaymentDiskon, Me.mUsePaymentDiskon) > 0 Then
                    Me.mPrimary = DTA.ScalarQuery_GetPrimary(Me.mCompanyName)
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary).Value
                    Me.mCreatedDate = Me.mTimestamp
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mCreatedBy = vUserCode
                    Me.mModifiedBy = vUserCode
                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, i.dgvAutorization, i)
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
            DTA_SAVE.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function SimpanPerubahan(ByVal vUserCode As String, _
    ByVal vObjCD As ClsCekData, ByVal vPrimary As Double, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Dim DTA_SAVE As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_SAVETableAdapter
        Dim ObjItem As New ClsTrans_POS_GeneralSetting_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahDataBenar(False, True, i.dgvAutorization, vEnumBhs, i) = True Then
                If DTA_SAVE.Update(Me.mCompanyName, Me.mShowCompanyName, Me.mShowCompanyLogo, Me.mCompanyLogo, _
                Me.mPrimaryCurr, Me.mAllowCreditCardPayment, Me.mAllowDebitCardPayment, Me.mAllowCashPayment, _
                Me.mAllowVoucherPayment, Me.mUseTaxRate, Me.mPersenTaxRate, Me.mAllowRounding, Me.mRoundingTipe, _
                Me.mRoundingValue, Me.mRoundingOn, Me.mPrimaryAcc, Me.mAllowToCancel, Me.mUseAuthorToCancel, _
                Me.mAllowUserToCancelWholeTrans, Me.mUseAuthorToCancelWholeTrans, Me.mAllowUserToModifyTrans, _
                Me.mUseAuthorToModifyTrans, Me.mAllowUserToModifyUnitPrice, Me.mUseAuthorToModifyUnitPrice, _
                Me.mAllowUserToModifyDiscountRate, Me.mUseAuthorToModifyDiscountRate, Me.mAllowUserToReprintReceipt, _
                Me.mUseAuthorToReprintReceipt, vUserCode, Now, Now, Me.mPrimaryPayment, Me.mAutomaticTransfer, _
                Me.mExpiryASC, Me.mBatchASC, Me.mAllowLocationControl, Me.mAllowBatchNumberControl, Me.mAllowExpiryControl, _
                Me.mInputLocationOnCashierForm, Me.mDefPaymentDiskon, Me.mUsePaymentDiskon, Me.mPrimary, Me.mTimestamp) > 0 Then
                    Me.mTimestamp = DTA.ScalarQuery_GetTimestamp(Me.mPrimary)
                    Me.mModifiedDate = Me.mTimestamp
                    Me.mModifiedBy = vUserCode
                    'hapus data item lama
                    ObjItem.HapusItem(Me.mPrimary)
                    'simpan data item
                    ObjItem.SimpanItem(Me.mPrimary, i.dgvAutorization, i)
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            DTA_SAVE.Dispose()
            ObjItem = Nothing
        End Try
    End Function

    Public Overridable Function Hapus(ByVal vPrimary As Double, ByVal vUserCode As String, _
    ByVal vStrictDelete As Boolean, ByVal vShowMsg As Boolean, _
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
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
    ByVal vEnumBhs As EnumBahasa, ByVal i As FrmTrans_POS_Settings) As Boolean
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Dim ObjItem As New ClsTrans_POS_GeneralSetting_Item
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If Me.ApakahBolehHapus(vPrimary, vUserCode, _
                vStrictDelete, vShowMsg, vEnumBhs, i) = True Then
                'Hapus item
                ObjItem.HapusItem(vPrimary)
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
            ObjItem = Nothing
        End Try
    End Function

    Public Sub ResetObject()
        Try
            Me.mPrimary = "-1"
            Me.mCompanyName = ""
            Me.mCompanyLogo = ""
            Me.mShowCompanyName = 0
            Me.mShowCompanyLogo = 0
            Me.mPrimaryCurr = "-1"
            Me.mAllowCreditCardPayment = 0
            Me.mAllowDebitCardPayment = 0
            Me.mAllowCashPayment = 0
            Me.mAllowVoucherPayment = 0
            Me.mUseTaxRate = 0
            Me.mPersenTaxRate = 0
            Me.mAllowRounding = 0
            Me.mRoundingTipe = EnumRoundingTipe.Up
            Me.mRoundingValue = EnumRoundingValue.G_Satu
            Me.mRoundingOn = EnumRoundingOn.Before_Tax_Calculation
            Me.mPrimaryAcc = "-1"
            Me.mAllowToCancel = 0
            Me.mUseAuthorToCancel = 0
            Me.mAllowUserToCancelWholeTrans = 0
            Me.mUseAuthorToCancelWholeTrans = 0
            Me.mAllowUserToModifyTrans = 0
            Me.mUseAuthorToModifyTrans = 0
            Me.mAllowUserToModifyUnitPrice = 0
            Me.mUseAuthorToModifyUnitPrice = 0
            Me.mAllowUserToModifyDiscountRate = 0
            Me.mUseAuthorToModifyDiscountRate = 0
            Me.mAllowUserToReprintReceipt = 0
            Me.mUseAuthorToReprintReceipt = 0
            Me.mAutomaticTransfer = EnumAutomaticTransfer.Stop_Transfer
            Me.mExpiryASC = EnumExpiryASC.Descending
            Me.mBatchASC = EnumBatchASC.Descending
            Me.mAllowLocationControl = 0
            Me.mAllowBatchNumberControl = 0
            Me.mAllowExpiryControl = 0
            Me.mInputLocationOnCashierForm = EnumInputLocationOnCashierForm.No

            mCreatedBy = ""
            mCreatedDate = Today
            mModifiedBy = ""
            mModifiedDate = Today
            mTimestamp = Today
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetJmlComapany() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetJml().HasValue Then
                Return DTA.ScalarQuery_GetJml().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCompany() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryCompany().HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCompany().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurr() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryCurr().HasValue Then
                Return DTA.ScalarQuery_GetPrimaryCurr().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAllowRounding() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetAllowRounding().HasValue Then
                Return DTA.ScalarQuery_GetAllowRounding().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRoundingTipe() As EnumRoundingTipe
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_RoundingTipe().HasValue Then
                Return DTA.ScalarQuery_RoundingTipe().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetRoundingvalue() As EnumRoundingValue
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetRoundingValue().HasValue Then
                Return DTA.ScalarQuery_GetRoundingValue().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUseTax() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetUseTaxRate().HasValue Then
                Return DTA.ScalarQuery_GetUseTaxRate().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPersenTax() As Decimal
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPersenTaxRate().HasValue Then
                Return DTA.ScalarQuery_GetPersenTaxRate().Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryPayment() As Double
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetPrimaryPayment().HasValue Then
                Return DTA.ScalarQuery_GetPrimaryPayment().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPaymentDiskon() As String
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetPaymentDiscount()) = False Then
                Return DTA.ScalarQuery_GetPaymentDiscount()
            Else
                Return "0"
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return "0"
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAllowToCancelItem() As Integer
        Dim dta As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            dta.Connection.ConnectionString = Me.GetConstring
            If dta.ScalarQuery_AllowToCancelItem().HasValue Then
                Return dta.ScalarQuery_AllowToCancelItem.Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiCBOConfirmPassword(ByVal vCBO As ComboBox, _
    ByVal vDTA As DseTrans_POS_GeneralSettingTableAdapters.tbluserTableAdapter, _
    ByVal vTBL As DseTrans_POS_GeneralSetting.tbluserDataTable, ByVal vTipeData As FrmTrans_POS_ConfirmPassword.EnumTipe)
        Dim DTAItem As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_setting_itemTableAdapter
        Dim TBLItem As New DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_itemDataTable
        Dim RowItem As DseTrans_POS_GeneralSetting.tbltrans_pos_general_setting_itemRow
        Dim DTADummy As New DseTrans_POS_GeneralSettingTableAdapters.tbluserTableAdapter
        Dim TBLDummy As New DseTrans_POS_GeneralSetting.tbluserDataTable
        Dim RowDummy As DseTrans_POS_GeneralSetting.tbluserRow
        Dim Row As DseTrans_POS_GeneralSetting.tbluserRow
        Dim ListKode As New List(Of String)
        Dim ParamKode As String = ""
        Dim i As Integer = 0
        Try
            vDTA.Connection.ConnectionString = Me.GetConstring
            DTAItem.Connection.ConnectionString = Me.GetConstring
            DTADummy.Connection.ConnectionString = Me.GetConstring

            vDTA.Fill(vTBL)
            DTADummy.Fill(TBLDummy)
            If vTipeData = FrmTrans_POS_ConfirmPassword.EnumTipe.ItemDeletion Then
                DTAItem.FillBy_PrimaryNItemDelete(TBLItem)
            ElseIf vTipeData = FrmTrans_POS_ConfirmPassword.EnumTipe.WholeTransactionDeletion Then
                DTAItem.FillBy_WholeTransactionDeletion(TBLItem)
            End If


            For Each RowItem In TBLItem.Rows
                If ListKode.Contains(RowItem.User_Code) = False Then
                    ListKode.Add(RowItem.User_Code)
                End If
            Next

            For Each ParamKode In ListKode
                While i <= TBLDummy.Rows.Count - 1
                    RowDummy = TBLDummy.Rows(i)
                    If RowDummy.Kode_User = ParamKode Then
                        TBLDummy.Rows.RemoveAt(i)
                    Else
                        i += 1
                    End If
                End While
            Next

            i = 0
            Dim Ketemu As Boolean = False
            While i <= vTBL.Rows.Count - 1
                Row = vTBL.Rows(i)
                For Each RowDummy In TBLDummy.Rows
                    If Row.Kode_User = RowDummy.Kode_User Then
                        vTBL.Rows.RemoveAt(i)
                        Ketemu = True
                        Exit For
                    End If
                Next
                If Ketemu = False Then
                    i += 1
                Else
                    Ketemu = False
                End If
            End While

            'kosongkan cbo
            vCBO.DataSource = Nothing
            vCBO.Items.Clear()

            'hubungkan dengan cbo
            For Each Row In vTBL.Rows
                vCBO.Items.Add(Row.Kode_User)
            Next
            vCBO.Focus()
            'vCBO.DataSource = vTBL.DefaultView
            'vTBL.DefaultView.Sort = "[Kode User] ASC"
            'vCBO.DisplayMember = "Kode User"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTAItem.Dispose()
            TBLItem.Dispose()
            DTADummy.Dispose()
            TBLDummy.Dispose()
            ListKode = Nothing
        End Try
    End Sub

    Public Function GetUseAuthorAllowCancelItem() As Integer
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetUseAuthorCancelItem().HasValue Then
                Return DTA.ScalarQuery_GetUseAuthorCancelItem().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAllowUserToCancelWholeTransaction() As Integer
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetAllow_User_To_Cancel_Whole_Transaction().HasValue Then
                Return DTA.ScalarQuery_GetAllow_User_To_Cancel_Whole_Transaction().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUseAuthorizationToCancelWholeTransaction() As Integer
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetUseAuthorizationToCancelWholeTransaction().HasValue Then
                Return DTA.ScalarQuery_GetUseAuthorizationToCancelWholeTransaction().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetAllowUserToReprintReceipt() As Integer
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetAllowUserToReprintReceipt().HasValue Then
                Return DTA.ScalarQuery_GetAllowUserToReprintReceipt().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetUseAuthorizationToReprintReceipt() As Integer
        Dim DTA As New DseTrans_POS_GeneralSettingTableAdapters.tbltrans_pos_general_settingTableAdapter
        Try
            DTA.Connection.ConnectionString = Me.GetConstring
            If DTA.ScalarQuery_GetUseAuthorizationToReprintReceipt().HasValue Then
                Return DTA.ScalarQuery_GetUseAuthorizationToReprintReceipt().Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function
End Class
