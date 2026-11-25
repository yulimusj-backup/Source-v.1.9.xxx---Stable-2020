Option Explicit On
Option Strict Off

Public Class ClsMasterAkun
    Dim ObjCD As New Core_Module.ClsCekData

    Private mPrimary As Double = -1
    Private mAccountCode As String = ""
    Private mDescription As String = ""
    Private mAccountNature As String = ""
    Private mAktif As Integer = 1
    Private mHeader As Integer = 0
    Private mBeginningBalance As Decimal = 0
    Private mTipePerkiraan As EnumTipeAkun = EnumTipeAkun.Asset
    Private mLevelPerkiraan As Integer = 1
    Private mTampilan As String = ""
    Private mSaldoNormal As EnumSaldoNormal = EnumSaldoNormal.Debet
    Private mUrutan As Integer = 0
    Private mPrimaryLv1 As Double = -1
    Private mPrimaryLv2 As Double = -1
    Private mPrimaryLv3 As Double = -1
    Private mTipeAkun2 As Integer = 0
    Private mBeginningBalanceOriginalCurrency As Decimal = 0
    Private mPrimaryCurrency As Double = -1
    Private mPrimaryLvl4 As Double = -1
    Private mPrimaryLvl5 As Double = -1
    Private mPrimaryLvl6 As Double = -1
    Private mPrimaryLvl7 As Double = -1

    Private mCreatedBy As String = ""
    Private mCreatedDate As Date = Now
    Private mModifiedBy As String = ""
    Private mModifiedDate As Date = Now
    Private mTimestamp As Date = Now

    Public Property PrimaryLvl4() As Double
        Get
            Return Me.mPrimaryLvl4
        End Get
        Set(ByVal value As Double)
            mPrimaryLvl4 = value
        End Set
    End Property

    Public Property PrimaryLvl5() As Double
        Get
            Return Me.mPrimaryLvl5
        End Get
        Set(ByVal value As Double)
            mPrimaryLvl5 = value
        End Set
    End Property

    Public Property PrimaryLvl6() As Double
        Get
            Return Me.mPrimaryLvl6
        End Get
        Set(ByVal value As Double)
            mPrimaryLvl6 = value
        End Set
    End Property

    Public Property PrimaryLvl7() As Double
        Get
            Return Me.mPrimaryLvl7
        End Get
        Set(ByVal value As Double)
            mPrimaryLvl7 = value
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

    Public Property SaldoNormal() As EnumSaldoNormal
        Get
            Return mSaldoNormal
        End Get
        Set(ByVal value As EnumSaldoNormal)
            mSaldoNormal = value
        End Set
    End Property

    Public Enum EnumSaldoNormal As Integer
        Debet
        Kredit
    End Enum

    Public Property TipePerkiraan() As EnumTipeAkun
        Get
            Return mTipePerkiraan
        End Get
        Set(ByVal value As EnumTipeAkun)
            mTipePerkiraan = value
        End Set
    End Property

    Public Enum EnumTipeAkun As Integer
        Asset
        Liabilities
        Equity
        Income
        Expense
        Cost_of_Goods_Sales
        Other_Income
        Other_Expense
        Tax
    End Enum

    Public Enum EnumTipeAkun2 As Integer
        Cash_And_Bank
        Account_Receivable
        Other_Receivable
        Inventory
        Fixed_Asset
        Accumulated_Depreciation
        Other_Asset
        Account_Payable
        Other_Current_Liability
        Long_Term_Liability
        Equity
        Income
        Cost_Of_Sales
        Expense
        Other_Income
        Other_Expense
        Tax
        Asset
        Liabilities
        Other_Payable
    End Enum

    Public Property TipeAkun2() As EnumTipeAkun2
        Get
            Return Me.mTipeAkun2
        End Get
        Set(ByVal value As EnumTipeAkun2)
            mTipeAkun2 = value
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

    Public Property Timestamp() As Date
        Get
            Return mTimestamp
        End Get
        Set(ByVal value As Date)
            mTimestamp = value
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

    Public Property AccountCode() As String
        Get
            Return mAccountCode
        End Get
        Set(ByVal value As String)
            mAccountCode = value
        End Set
    End Property

    Public Property Description() As String
        Get
            Return mDescription
        End Get
        Set(ByVal value As String)
            mDescription = value
        End Set
    End Property

    Public Property AccountNature() As String
        Get
            Return mAccountNature
        End Get
        Set(ByVal value As String)
            mAccountNature = value
        End Set
    End Property

    Public Property Aktif() As Integer
        Get
            Return mAktif
        End Get
        Set(ByVal value As Integer)
            mAktif = value
        End Set
    End Property

    Public Property Header() As Integer
        Get
            Return mHeader
        End Get
        Set(ByVal value As Integer)
            mHeader = value
        End Set
    End Property

    Public Property BeginningBalanceOriginalCurrency() As Decimal
        Get
            Return Me.mBeginningBalanceOriginalCurrency
        End Get
        Set(ByVal value As Decimal)
            mBeginningBalanceOriginalCurrency = value
        End Set
    End Property

    Public Property BeginningBalance() As Decimal
        Get
            Return mBeginningBalance
        End Get
        Set(ByVal value As Decimal)
            mBeginningBalance = value
        End Set
    End Property

    Public Property Tampilan() As String
        Get
            Return mTampilan
        End Get
        Set(ByVal value As String)
            mTampilan = value
        End Set
    End Property

    Public Property LevelPerkiraan() As Integer
        Get
            Return mLevelPerkiraan
        End Get
        Set(ByVal value As Integer)
            mLevelPerkiraan = value
        End Set
    End Property

    Public Property Urutan() As Integer
        Get
            Return mUrutan
        End Get
        Set(ByVal value As Integer)
            mUrutan = value
        End Set
    End Property

    Public Property PrimaryLv1() As Double
        Get
            Return mPrimaryLv1
        End Get
        Set(ByVal value As Double)
            mPrimaryLv1 = value
        End Set
    End Property

    Public Property PrimaryLv2() As Double
        Get
            Return mPrimaryLv2
        End Get
        Set(ByVal value As Double)
            mPrimaryLv2 = value
        End Set
    End Property

    Public Property PrimaryLv3() As Double
        Get
            Return mPrimaryLv3
        End Get
        Set(ByVal value As Double)
            mPrimaryLv3 = value
        End Set
    End Property

    Public Function GetTimestamp(ByVal vPrimary As Double) As Date
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTimestamp(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetTimestamp(vPrimary)
            Else
                Return Date.MinValue
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetBalance_AkunNonHeader(ByVal vPrimaryAkun As Double, ByVal vTanggalSampai As Date) As Decimal
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Dim TBL As New DseMasterAkun.tblmaster_akunDataTable
        Dim row As DseMasterAkun.tblmaster_akunRow
        Dim ObjClosing As New ClsTransaksiClosing

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ObjClosing.GetConvertionDate.Date > vTanggalSampai.Date + TimeSpan.FromDays(1) Then
                Return 0
            Else
                If DTA.FillBy_PrimaryAkun_TanggalSampai(TBL, vPrimaryAkun, vTanggalSampai) > 0 Then
                    row = TBL.Rows(0)
                    If row.Header = 1 Then
                        'fungsi ini hanya terbatas untuk akun2 non header
                        Return 0
                    Else
                        If row.Saldo_Normal = 0 Then
                            'debet
                            Return row.Beginning_Balance + row.Debet - row.Kredit
                        Else
                            'credit
                            Return row.Beginning_Balance - row.Debet + row.Kredit
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetBalanceOriginal_AkunNonHeader(ByVal vPrimaryAkun As Double, ByVal vTanggalSampai As Date) As Decimal
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun5TableAdapter
        Dim TBL As New DseMasterAkun.tblmaster_akun5DataTable
        Dim row As DseMasterAkun.tblmaster_akun5Row
        Dim ObjClosing As New ClsTransaksiClosing

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If ObjClosing.GetConvertionDate.Date > vTanggalSampai.Date + TimeSpan.FromDays(1) Then
                Return 0
            Else
                If DTA.Fill(TBL, vPrimaryAkun, vTanggalSampai) > 0 Then
                    row = TBL.Rows(0)
                    If row.Header = 1 Then
                        'fungsi ini hanya terbatas untuk akun2 non header
                        Return 0
                    Else
                        If row.Saldo_Normal = 0 Then
                            'debet
                            Return row.Beggining_Balance_Original_Cur + row.Debet_Original - row.Credit_Original
                        Else
                            'credit
                            Return row.Beggining_Balance_Original_Cur - row.Debet_Original + row.Credit_Original
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL.Dispose()
            ObjClosing = Nothing
        End Try
    End Function

    Public Function GetAccountCode(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetAccountCodeByPrimary(vPrimary)
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

    Public Function GetTampilan(ByVal vPrimary As Double) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetTampilanByPrimary(vPrimary)) = False Then
                Return DTA.ScalarQuery_GetTampilanByPrimary(vPrimary)
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

    Public Function GetJmlAnak(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Dim Lvl As Integer = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek level
            If DTA.ScalarQuery_GetLevel(vPrimaryAkun).HasValue = True Then
                Lvl = DTA.ScalarQuery_GetLevel(vPrimaryAkun).Value
            Else
                Return 0
            End If

            If DTA.ScalarQuery_GetLevel(vPrimaryAkun).HasValue = True Then
                If Lvl = 1 Then
                    'jika level 1
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl1(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl1(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 2 Then
                    'jika level 2
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl2(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl2(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 3 Then
                    'jika level 3
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl3(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl3(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 4 Then
                    'jika level 4
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl4(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl4(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 5 Then
                    'jika level 5
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl5(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl5(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 6 Then
                    'jika level 6
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl6(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl6(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 7 Then
                    'jika level 7
                    If DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl7(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAnak_ByPrimaryLvl7(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                Else
                    'jika level 8, ga mungkin punya anak
                    Return 0
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetSaldoNormal(ByVal vPrimary As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetSaldoNormal(vPrimary).HasValue Then
                Return DTA.ScalarQuery_GetSaldoNormal(vPrimary).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetPrimaryCurrency(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetPrimaryCurrency(vPrimaryAkun).Value
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

    Public Function GetPrimary(ByVal vAccountCode As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimary(vAccountCode).HasValue Then
                Return DTA.ScalarQuery_GetPrimary(vAccountCode).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Overridable Function ApakahBolehEdit(ByVal vPrimaryTransaksi As Double, _
         Optional ByVal TampilkanPesan As Boolean = True) As Boolean
        Try
            'sampe sini balikan true
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally

        End Try
    End Function

    Public Function ApakahBolehHapus(ByVal vPrimaryTransaksi As Double, _
       Optional ByVal TampilkanPesan As Boolean = True, Optional ByVal vCekCurrency As Boolean = True) As Boolean
        Dim ObjCurr As New Core_Module.clsMasterCurrency
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_productTableAdapter
        Dim TBL As New DseMasterAkun.tblmaster_productDataTable
        Dim row As DseMasterAkun.tblmaster_productRow
        Dim Dummy As String = ""
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim DTATipeSettingAdj As New DseMasterAkunTableAdapters.tbl_setting_adjustmentTableAdapter
        Dim TBLTipeSettingAdj As New DseMasterAkun.tbl_setting_adjustmentDataTable
        Dim RowTipeSettingAdj As DseMasterAkun.tbl_setting_adjustmentRow
        Dim DTAClosing As New DseMasterAkunTableAdapters.tbltrans_purchase_closingTableAdapter
        Dim TBLClosing As New DseMasterAkun.tbltrans_purchase_closingDataTable
        Dim rowClosing As DseMasterAkun.tbltrans_purchase_closingRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTATipeSettingAdj.Connection.ConnectionString = DTA.Connection.ConnectionString
            DTAClosing.Connection.ConnectionString = DTA.Connection.ConnectionString

            'data Proses boleh dihapus jika
            '1. Akun dengan primary 1-8 tidak boleh dihapus
            If vPrimaryTransaksi <= 9 Then
                If TampilkanPesan = True Then
                    MsgBox("This is level 1 account. You can't delete this account.", MsgBoxStyle.Information, "Information")
                End If

                Return False
            End If
            '2. Untuk akun yang memiliki anak tidak boleh di hapus
            If Me.GetJmlAnak(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used by the other account", MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            '3. Sudah ada item jurnal yang merujuk ke data tersebut
            Dim ObjJurnal As New ClsTransaksiJournal
            If ObjJurnal.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in journal, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjJurnal = Nothing

            '4. Sudah ada item budget yang merujuk ke data tersebut (Confirm, kalo ga keberatan maka item
            '   budget tersebut dihapus dahulu baru hapus data akun)
            Dim Obj_Budget As New ClsMasterBudget
            If Obj_Budget.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in budget data, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            Obj_Budget = Nothing

            '5. Sudah digunakan dalam transaksi invoice
            Dim ObjSett_Inv As New ClsTransaksiSettingTransf_Inv
            If ObjSett_Inv.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in invoice data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_Inv = Nothing

            '6. Sudah digunakan dalam transaksi invoice DP
            Dim ObjSett_InvDP As New ClsTransaksiSettingTransf_InvDP
            If ObjSett_InvDP.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in downpayment invoice data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_InvDP = Nothing

            '7. Sudah digunakan dalam transaksi sales return
            Dim ObjSett_SR As New ClsTransaksiSettingTransf_SR
            If ObjSett_SR.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in sales return data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_SR = Nothing

            '8. Sudah digunakan dalam transaksi AOP
            Dim ObjSett_AOP As New ClsTransaksiSettingTransf_AOP
            If ObjSett_AOP.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in AOP data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_AOP = Nothing

            '9. Sudah digunakan dalam transaksi AOP-DP
            Dim ObjSett_AOP_DP As New ClsTransaksiSettingTransf_AOPDP
            If ObjSett_AOP_DP.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in downpayment AOP data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_AOP_DP = Nothing

            '10. Sudah digunakan dalam transaksi AOPR
            Dim ObjSett_AOPR As New ClsTransaksiSettingTrans_AOPR
            If ObjSett_AOPR.GetJml_ByPrimaryAkun(vPrimaryTransaksi) > 0 Then
                If TampilkanPesan = True Then
                    MsgBox("This account is used in downpayment AOPR data transfer, you can not delete this account.", _
                    MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If
            ObjSett_AOPR = Nothing

            If vCekCurrency = True Then
                '11. tidak boleh dihapus jika sudah dirujuk oleh master currency oleh account payable
                If ObjCurr.GetJmlByPrimaryAcc_AccountPayable(vPrimaryTransaksi) > 0 Then
                    MsgBox("This account refers by currency data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '12. tidak boleh dihapus jika sudah dirujuk oleh master currency oleh account receivable
                If ObjCurr.GetJmlByPrimaryAcc_AccountReceivable(vPrimaryTransaksi) > 0 Then
                    MsgBox("This account refers by currency data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '13. tidak boleh dihapus jika sudah duirujuk oleh master currency oleh unrealized exchange gain/loss account
                If ObjCurr.GetJmlByPrimaryAcc_UnrealizedExchangeGainLossAccount(vPrimaryTransaksi) > 0 Then
                    MsgBox("This account refers by currency data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
                '14. tidak boleh dihapus jika sudah dirujuk oleh master currency oleh realized exchange gain/loss account
                If ObjCurr.GetJmlByPrimaryAcc_RealizedExchangeGainLossAccount(vPrimaryTransaksi) > 0 Then
                    MsgBox("This account refers by currency data, you cannot delete this data.", _
                    MsgBoxStyle.Information, "Information")
                    Return False
                End If
            End If

            'jika sudah digunakan dalam product
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(TBL, vPrimaryTransaksi)
            If TBL.Rows.Count > 0 Then
                For Each row In TBL.Rows
                    If Dummy = "" Then
                        Dummy = row.Product_Code
                    Else
                        Dummy = Dummy + ", " + row.Product_Code
                    End If
                Next
                MsgBox("The following product(s) refer to this account, you cannot delete this data." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                Return False
            End If

            'jika sudah digunakan dalam setting global
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_CurrenctYearEarning)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Current year earning setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            'retained earning
            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Retained_Earning)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Retained earning setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_In)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("VAT In setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_VAT_Out)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("VAT Out setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Dutystamp_Expense)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Duty stamp expense setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_PaymentDifference)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Payment difference expense setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_ShippingIncome)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Shipping income setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_AOP_DP)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Purchase downpayment account setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            Dummy = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Jurnal, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.Akun_Invoice_DP)
            If Dummy = vPrimaryTransaksi.ToString Then
                MsgBox("Sales downpayment account setting in global setting refer to this account, you cannot delete this data.", _
                MsgBoxStyle.Information, "Information")
                Return False
            End If

            DTATipeSettingAdj.Fill(TBLTipeSettingAdj, vPrimaryTransaksi)
            If TBLTipeSettingAdj.Rows.Count > 0 Then
                Dummy = ""
                For km As Integer = 0 To TBLTipeSettingAdj.Rows.Count - 1
                    RowTipeSettingAdj = TBLTipeSettingAdj.Rows(km)
                    If Dummy = "" Then
                        Dummy = (km + 1).ToString + ". " + RowTipeSettingAdj.Code + " - " + RowTipeSettingAdj.Desc
                    Else
                        Dummy = Dummy + vbCrLf + (km + 1).ToString + ". " + RowTipeSettingAdj.Code + " - " + RowTipeSettingAdj.Desc
                    End If
                Next

                If TampilkanPesan = True Then
                    MsgBox("The following adjustment type is refering to this account, you cannot delete this data." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'closing, terkait jurnal yang digunakan
            DTAClosing.Fill(TBLClosing, vPrimaryTransaksi)
            If TBLClosing.Rows.Count > 1 Then
                Dummy = ""

                For km As Integer = 0 To TBLClosing.Rows.Count - 1
                    rowClosing = TBLClosing.Rows(km)
                    If Dummy = "" Then
                        Dummy = (km + 1).ToString + ". " + rowClosing.No_Transaksi
                    Else
                        Dummy = Dummy + vbCrLf + (km + 1).ToString + ". " + rowClosing.No_Transaksi
                    End If
                Next

                If TampilkanPesan = True Then
                    MsgBox("The following monthly closing refers to this account, you cannot delete this data." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            'closing, terkait overhead
            DTAClosing.FillBy1(TBLClosing, vPrimaryTransaksi)
            If TBLClosing.Rows.Count > 1 Then
                Dummy = ""

                For km As Integer = 0 To TBLClosing.Rows.Count - 1
                    rowClosing = TBLClosing.Rows(km)
                    If Dummy = "" Then
                        Dummy = (km + 1).ToString + ". " + rowClosing.No_Transaksi
                    Else
                        Dummy = Dummy + vbCrLf + (km + 1).ToString + ". " + rowClosing.No_Transaksi
                    End If
                Next

                If TampilkanPesan = True Then
                    MsgBox("The following monthly closing refers to this account, you cannot delete this data." + vbCrLf + Dummy, MsgBoxStyle.Information, "Information")
                End If
                Return False
            End If

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            ObjCurr = Nothing
            DTA.Dispose()
            TBL.Dispose()
            ObjSetting = Nothing
            DTATipeSettingAdj.Dispose()
            TBLTipeSettingAdj.Dispose()
            DTAClosing.Dispose()
            TBLClosing.Dispose()
        End Try
    End Function

    Public Overridable Function ApakahDataBenar(ByVal vBaru As Boolean) As Boolean
        Try
            If vBaru = True Then
                If Me.ApakahSudahAda(Me.mAccountCode) = True Then
                    MsgBox("Sudah ada akun dengan no " + Me.mAccountCode + ", ganti akun code.", MsgBoxStyle.Information, "Information")
                    Return False
                End If
            Else
                'no po tidak boleh duplikat
                Dim NoLama As String = Me.GetAccountCode(Me.mPrimary)
                If Me.AccountCode <> NoLama Then
                    If Me.ApakahSudahAda(Me.mAccountCode) = True Then
                        MsgBox("Sudah ada akun dengan no " + Me.mAccountCode + ", ganti akun code.", MsgBoxStyle.Information, "Information")
                        Return False
                    End If
                End If
            End If

            'return value
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function ApakahSudahAda(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJmlByAccountCode(vAccountCode).Value > 0 Then
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

    Public Function ApakahSudahAda_Aktif_NonHeader(ByVal vAccountCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).HasValue Then
                If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vAccountCode).Value > 0 Then
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

    Public Function ApakahSudahAda_Aktif_NonHeader2(ByVal vPrimaryAkun As Double) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJmlAktif_NonHeader(vPrimaryAkun).HasValue Then
                If DTA.ScalarQuery_GetJmlAktif_NonHeader(vPrimaryAkun).Value > 0 Then
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

    Public Function SimpanBaru(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter
        Dim TBL As New DseMasterAkun.tblmaster_akun1DataTable
        Dim row As DseMasterAkun.tblmaster_akun1Row
        Dim Start As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.Insert(Me.AccountCode, Me.Description, Me.AccountNature, _
                Me.Aktif, Me.Header, Me.BeginningBalance, Me.TipePerkiraan, Me.LevelPerkiraan, _
                Me.Tampilan, Me.SaldoNormal, Me.Urutan, Me.PrimaryLv1, Me.PrimaryLv2, Me.PrimaryLv3, _
                vUserCode, Now, vUserCode, Now, Now, Me.mPrimaryCurrency, Me.mBeginningBalanceOriginalCurrency, Me.mTipeAkun2, _
                Me.mPrimaryLvl4, Me.mPrimaryLvl5, Me.mPrimaryLvl6, Me.mPrimaryLvl7) > 0 Then
                Me.mPrimary = Me.GetPrimary(Me.mAccountCode)
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mCreatedDate = Me.mTimestamp
                Me.mModifiedDate = Me.mTimestamp
                Me.mCreatedBy = vUserCode
                Me.mModifiedBy = vUserCode

                'update data urutan
                DTA.FillBy_Urutan(TBL, Me.mUrutan, Me.mPrimary)
                Start = Me.Urutan + 1
                For Each row In TBL.Rows
                    DTA.UpdateQuery_Urutan(Start, row.Primary)
                    Start += 1
                Next

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

    Public Sub IsikanCLBAkun(ByVal vTBL As DseMasterAkun.tblmaster_akun7DataTable, ByVal vCLB As Windows.Forms.CheckedListBox)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun7TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            vCLB.DataSource = Nothing
            vCLB.Items.Clear()

            DTA.Fill(vTBL)
            vCLB.DataSource = vTBL.DefaultView
            vCLB.DisplayMember = "Account"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Sub IsiCBOBeban(ByVal vCBO As ComboBox, ByVal vTBL As DseMasterAkun.tblmaster_akun3DataTable)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun3TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL)
            vTBL.DefaultView.Sort = "Account"
            vCBO.DataSource = vTBL.DefaultView
            vCBO.DisplayMember = "Account"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Function GetPrimaryByDesc(ByVal vAccountDesc As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun4TableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetPrimaryByDesc(vAccountDesc).HasValue Then
                Return DTA.ScalarQuery_GetPrimaryByDesc(vAccountDesc).Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Sub IsiCboCashAndBank(ByVal vCbo As ComboBox, ByVal vTBL As DseMasterAkun.tblmaster_akun4DataTable)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun4TableAdapter

        Try
            vCbo.DataSource = Nothing
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_TypeAkun(vTBL, ClsMasterAkun.EnumTipeAkun.Asset, ClsMasterAkun.EnumTipeAkun2.Cash_And_Bank)

            vTBL.DefaultView.Sort = "Description"
            vCbo.DataSource = vTBL.DefaultView
            vCbo.DisplayMember = "Description"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCboPayment(ByVal vPrimaryCurrency As Double, ByVal vCbo As ComboBox, ByVal vTBL As DseMasterAkun.tblmaster_akun4DataTable)
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun4TableAdapter
        Dim TipePerkiraan2 As EnumTipeAkun2 = EnumTipeAkun2.Cash_And_Bank
        Dim TipePerkiraan1 As EnumTipeAkun = EnumTipeAkun.Asset

        Try
            vCbo.DataSource = Nothing
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.Fill(vTBL, vPrimaryCurrency, Core_Module.PrimaryBaseCurrency, TipePerkiraan2, TipePerkiraan1)

            vTBL.DefaultView.Sort = "Description"
            vCbo.DataSource = vTBL.DefaultView
            vCbo.DisplayMember = "Description"
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Sub

    Public Sub IsiCBOTipeAkun2(ByVal vTipeAkun As EnumTipeAkun, ByVal vCBO As ComboBox, ByVal vLvlAkun As Integer)
        Try
            vCBO.Items.Clear()

            If vTipeAkun = EnumTipeAkun.Asset Then
                If vLvlAkun = 1 Then
                    vCBO.Items.Add("Asset")
                Else
                    vCBO.Items.Add("Asset")
                    vCBO.Items.Add("Cash And Bank")
                    vCBO.Items.Add("Account Receivable")
                    vCBO.Items.Add("Other Receivable")
                    vCBO.Items.Add("Inventory")
                    vCBO.Items.Add("Fixed Asset")
                    vCBO.Items.Add("Accumulated Depreciation")
                    vCBO.Items.Add("Other Asset")
                End If
            ElseIf vTipeAkun = EnumTipeAkun.Liabilities Then
                If vLvlAkun = 1 Then
                    vCBO.Items.Add("Liabilities")
                Else
                    vCBO.Items.Add("Account Payable")
                    vCBO.Items.Add("Other Payable")
                    vCBO.Items.Add("Other Current Liability")
                    vCBO.Items.Add("Long Term Liability")
                End If
            ElseIf vTipeAkun = EnumTipeAkun.Equity Then
                vCBO.Items.Add("Equity")
            ElseIf vTipeAkun = EnumTipeAkun.Income Then
                vCBO.Items.Add("Income")
            ElseIf vTipeAkun = EnumTipeAkun.Cost_of_Goods_Sales Then
                vCBO.Items.Add("Cost Of Sales")
            ElseIf vTipeAkun = EnumTipeAkun.Expense Then
                vCBO.Items.Add("Expense")
            ElseIf vTipeAkun = EnumTipeAkun.Other_Income Then
                vCBO.Items.Add("Other Income")
            ElseIf vTipeAkun = EnumTipeAkun.Other_Expense Then
                vCBO.Items.Add("Other Expense")
            ElseIf vTipeAkun = EnumTipeAkun.Tax Then
                vCBO.Items.Add("Tax")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Public Function GetTipeAkun2(ByVal vTipeAkun2Str As String) As EnumTipeAkun2
        Try
            If vTipeAkun2Str = "Cash And Bank" Then
                Return EnumTipeAkun2.Cash_And_Bank
            ElseIf vTipeAkun2Str = "Account Receivable" Then
                Return EnumTipeAkun2.Account_Receivable
            ElseIf vTipeAkun2Str = "Other Receivable" Then
                Return EnumTipeAkun2.Other_Receivable
            ElseIf vTipeAkun2Str = "Inventory" Then
                Return EnumTipeAkun2.Inventory
            ElseIf vTipeAkun2Str = "Fixed Asset" Then
                Return EnumTipeAkun2.Fixed_Asset
            ElseIf vTipeAkun2Str = "Accumulated Depreciation" Then
                Return EnumTipeAkun2.Accumulated_Depreciation
            ElseIf vTipeAkun2Str = "Other Asset" Then
                Return EnumTipeAkun2.Other_Asset
            ElseIf vTipeAkun2Str = "Account Payable" Then
                Return EnumTipeAkun2.Account_Payable
            ElseIf vTipeAkun2Str = "Other Current Liability" Then
                Return EnumTipeAkun2.Other_Current_Liability
            ElseIf vTipeAkun2Str = "Long Term Liability" Then
                Return EnumTipeAkun2.Long_Term_Liability
            ElseIf vTipeAkun2Str = "Equity" Then
                Return EnumTipeAkun2.Equity
            ElseIf vTipeAkun2Str = "Income" Then
                Return EnumTipeAkun2.Income
            ElseIf vTipeAkun2Str = "Cost Of Sales" Then
                Return EnumTipeAkun2.Cost_Of_Sales
            ElseIf vTipeAkun2Str = "Expense" Then
                Return EnumTipeAkun2.Expense
            ElseIf vTipeAkun2Str = "Other Income" Then
                Return EnumTipeAkun2.Other_Income
            ElseIf vTipeAkun2Str = "Other Expense" Then
                Return EnumTipeAkun2.Other_Expense
            ElseIf vTipeAkun2Str = "Tax" Then
                Return EnumTipeAkun2.Tax
            ElseIf vTipeAkun2Str = "Asset" Then
                Return EnumTipeAkun2.Asset
            ElseIf vTipeAkun2Str = "Liabilities" Then
                Return EnumTipeAkun2.Liabilities
            ElseIf vTipeAkun2Str = "Other Payable" Then
                Return EnumTipeAkun2.Other_Payable
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeAkun2Default(ByVal vTipeAkun As EnumTipeAkun) As EnumTipeAkun2
        Try
            If vTipeAkun = EnumTipeAkun.Asset Then
                Return EnumTipeAkun2.Cash_And_Bank
            ElseIf vTipeAkun = EnumTipeAkun.Liabilities Then
                Return EnumTipeAkun2.Other_Current_Liability
            ElseIf vTipeAkun = EnumTipeAkun.Equity Then
                Return EnumTipeAkun2.Equity
            ElseIf vTipeAkun = EnumTipeAkun.Income Then
                Return EnumTipeAkun2.Income
            ElseIf vTipeAkun = EnumTipeAkun.Cost_of_Goods_Sales Then
                Return EnumTipeAkun2.Cost_Of_Sales
            ElseIf vTipeAkun = EnumTipeAkun.Expense Then
                Return EnumTipeAkun2.Expense
            ElseIf vTipeAkun = EnumTipeAkun.Other_Income Then
                Return EnumTipeAkun2.Other_Income
            ElseIf vTipeAkun = EnumTipeAkun.Other_Expense Then
                Return EnumTipeAkun2.Other_Expense
            ElseIf vTipeAkun = EnumTipeAkun.Tax Then
                Return EnumTipeAkun2.Tax
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function GetTipeAkun2(ByVal vTipeAkun2 As EnumTipeAkun2) As String
        Try
            If vTipeAkun2 = EnumTipeAkun2.Cash_And_Bank Then
                Return "Cash And Bank"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Account_Receivable Then
                Return "Account Receivable"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Receivable Then
                Return "Other Receivable"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Inventory Then
                Return "Inventory"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Fixed_Asset Then
                Return "Fixed Asset"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Accumulated_Depreciation Then
                Return "Accumulated Depreciation"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Asset Then
                Return "Other Asset"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Account_Payable Then
                Return "Account Payable"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Current_Liability Then
                Return "Other Current Liability"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Long_Term_Liability Then
                Return "Long Term Liability"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Equity Then
                Return "Equity"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Income Then
                Return "Income"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Cost_Of_Sales Then
                Return "Cost Of Sales"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Expense Then
                Return "Expense"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Income Then
                Return "Other Income"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Expense Then
                Return "Other Expense"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Tax Then
                Return "Tax"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Asset Then
                Return "Asset"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Liabilities Then
                Return "Liabilities"
            ElseIf vTipeAkun2 = EnumTipeAkun2.Other_Payable Then
                Return "Other Payable"
            End If

            Return ""
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        End Try
    End Function

    Public Function GetTipeAkun(ByVal vPrimary As Double) As EnumTipeAkun
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring

            If DTA.ScalarQuery_GetTipePerkiraan(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetTipePerkiraan(vPrimary).Value
            Else
                Return EnumTipeAkun.Asset
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function SimpanPerubahan(ByVal vUserCode As String) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.UpdateQuery_NoAccount(Me.AccountCode, Me.Description, Me.Tampilan, _
                vUserCode, Now, Now, Me.mAccountNature, Me.mAktif, Me.mHeader, Me.mBeginningBalance, Me.mPrimaryCurrency, _
                Me.mBeginningBalanceOriginalCurrency, Me.mTipeAkun2, Me.mPrimary, Me.mTimestamp) > 0 Then
                Me.mTimestamp = Me.GetTimestamp(Me.mPrimary)
                Me.mModifiedDate = Me.mTimestamp
                Me.mModifiedBy = vUserCode
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

    Public Function GetTipeAkun2ByPrimaryAkun(ByVal vPrimary As Double) As Double
        Dim DTA As New AccModule.DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Dim Dummy As System.Nullable(Of Long)

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Dummy = DTA.ScalarQuery_GetTipePerkiraan2(vPrimary)
            If Dummy.HasValue Then
                Return Dummy.Value
            Else
                Return -1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Private Function Get_Urutan_Akun(ByVal vPrimary As Double) As Integer
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetUrutan(vPrimary).HasValue = True Then
                Return DTA.ScalarQuery_GetUrutan(vPrimary).Value
            Else
                Return 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function Hapus(ByVal vPrimary As Double) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akun1TableAdapter
        Dim Start As Integer = 0
        Dim TBL As New DseMasterAkun.tblmaster_akun1DataTable
        Dim row As DseMasterAkun.tblmaster_akun1Row
        Dim Urutan As Integer = 0

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If Me.ApakahBolehHapus(vPrimary) = True Then
                Urutan = Me.Get_Urutan_Akun(vPrimary)

                DTA.Delete(vPrimary)

                'update data urutan
                DTA.FillBy_Urutan(TBL, Me.mUrutan, Me.mPrimary)

                Start = Me.Urutan
                For Each row In TBL.Rows
                    DTA.UpdateQuery_Urutan(Start, row.Primary)
                    Start += 1
                Next

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

    Public Function ApakahBolehGunakanForeignCurrency(ByVal vTipeAkun2 As EnumTipeAkun2) As Boolean
        Try
            If vTipeAkun2 = EnumTipeAkun2.Account_Payable Or _
            vTipeAkun2 = EnumTipeAkun2.Other_Payable Or _
            vTipeAkun2 = EnumTipeAkun2.Account_Receivable Or _
            vTipeAkun2 = EnumTipeAkun2.Other_Receivable Or _
            vTipeAkun2 = EnumTipeAkun2.Cash_And_Bank Or _
            vTipeAkun2 = EnumTipeAkun2.Other_Asset Or _
            vTipeAkun2 = EnumTipeAkun2.Long_Term_Liability Or _
            vTipeAkun2 = EnumTipeAkun2.Other_Current_Liability Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Function

    Public Function CariDetail(ByVal Primary As Double) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblCariDetail_master_akunTableAdapter
        Dim TBL As New DseMasterAkun.tblCariDetail_master_akunDataTable
        Dim row As DseMasterAkun.tblCariDetail_master_akunRow

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            DTA.FillBy_Primary(TBL, Primary)

            If TBL.Rows.Count > 0 Then
                row = TBL.Rows(0)
                With row
                    Me.mPrimary = .Primary
                    Me.mAccountCode = .Account_Code
                    Me.mDescription = .Description
                    Me.mAccountNature = .Account_Nature
                    Me.mAktif = .Aktif
                    Me.mHeader = .Header
                    Me.mBeginningBalance = .Beginning_Balance
                    Me.mTipePerkiraan = .Tipe_Perkiraan
                    Me.mLevelPerkiraan = .Level_Perkiraan
                    Me.mTampilan = .Tampilan
                    Me.mSaldoNormal = .Saldo_Normal
                    Me.mUrutan = .Urutan
                    Me.mPrimaryLv1 = .Primary_Lvl_1
                    Me.mPrimaryLv2 = .Primary_Lvl_2
                    Me.mPrimaryLv3 = .Primary_Lvl_3
                    Me.mPrimaryCurrency = .primary_currency
                    Me.mTipeAkun2 = .Tipe_Perkiraan2
                    Me.mBeginningBalanceOriginalCurrency = .Beggining_Balance_Original_Cur

                    Me.mCreatedBy = .Created_By
                    Me.mCreatedDate = .Created_Date
                    Me.mModifiedBy = .Modified_By
                    Me.mModifiedDate = .Modified_Date
                    Me.mTimestamp = .Timestamp
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

    Public Function GetJml(ByVal vKriteria As String, ByVal vSearchBy As ModGlobalVariable.EnumSearchCriteria) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If vSearchBy = EnumSearchCriteria.ByAccountNo Then
                Return DTA.ScalarQuery_GetJmlByAccountCodeLIKE(vKriteria + "%").Value
            ElseIf vSearchBy = EnumSearchCriteria.ByDescription Then
                Return DTA.ScalarQuery_GetJmlByDescLIKE(vKriteria + "%").Value
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function IsiDGV(ByVal Criteria As String, _
    ByVal DGV As DataGridView, ByVal TBL As DseMasterAkun.tbl_Acari_master_akunDataTable, _
    ByVal SearchCriteria As EnumSearchCriteria, ByVal tsl As ToolStripLabel) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tbl_Acari_master_akunTableAdapter
        Dim QueryResult As Double
        Dim StartT As Date

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            StartT = Now
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If SearchCriteria = EnumSearchCriteria.ByAccountNo Then
                QueryResult = DTA.FillBy_AccountCodeLIKE(TBL, Criteria + "%")
            ElseIf SearchCriteria = EnumSearchCriteria.ByDescription Then
                QueryResult = DTA.FillBy_DescLIKE(TBL, Criteria + "%")
            End If

            'update dgv
            DGV.DataSource = TBL.DefaultView
            DGV.Refresh()
            If DGV.ColumnCount > 0 Then
                DGV.Columns(0).Visible = False
                DGV.Columns(1).HeaderText = "Account Code"
                DGV.Columns(0).Frozen = True
                DGV.Columns(1).Frozen = True
                DGV.Columns(2).HeaderText = "Description"
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

            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            Windows.Forms.Cursor.Current = Cursors.Arrow
        End Try
    End Function

    Public Function CetakMasterAkun(ByVal CRV As CrystalDecisions.Windows.Forms.CrystalReportViewer) As Boolean
        Dim Report As New Crp_CetakMasterAkun
        Dim DS As New DseMasterAkun
        Dim FormatDay As Integer = 1
        Dim FormatMonth As Integer = 2
        Dim FormatYears As Integer = 1
        Dim FormatDecimal As Integer = 2
        Dim DTA As New DseMasterAkunTableAdapters.tblCetak_master_akunTableAdapter
        Dim TBL As DseMasterAkun.tblCetak_master_akunDataTable = DS.tblCetak_master_akun
        Dim row As DseMasterAkun.tblCetak_master_akunRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim ObjCur As New Core_Module.clsMasterCurrency
        Dim DefaultPrecision As Integer = 0
        Dim i As Integer = 0
        Dim ObjSetting As New Core_Module.ClsTransaksiSetting
        Dim Negatif As Integer = 0

        Dim j As EnumTipeAkun2 = EnumTipeAkun2.Account_Payable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            FormatDecimal = ObjCur.GetDecimal_DefaultCurrency
            'isikan tabel
            If DTA.Fill(TBL) > 0 Then
                For Each row In TBL.Rows
                    j = row.Tipe_Perkiraan2
                    row.Account_Type = Me.GetTipeAkun2(j)
                Next
                '    For Each row In TBL.Rows
                '        With row
                '            If .Header = 0 Then
                '                If .Saldo_Normal = 0 Then
                '                    .Kredit = 0
                '                    .Debet = Math.Round(.Beginning_Balance + .Debet - .Kredit, FormatDecimal, MidpointRounding.AwayFromZero)
                '                    .Tebal = 0
                '                Else
                '                    .Kredit = Math.Round(.Beginning_Balance + _
                '                    .Kredit - .Debet, FormatDecimal, MidpointRounding.AwayFromZero)
                '                    .Debet = 0
                '                    .Tebal = 0
                '                End If
                '            End If
                '        End With
                '    Next

                '    For Each row In TBL.Rows
                '        If row.Header = 1 Then
                '            If row.Saldo_Normal = 0 Then
                '                row.Debet = Me.GetSaldoHeaderCetakan(row.Primary, _
                '                row.Level_Perkiraan, TBL, row.Saldo_Normal)
                '                row.Tebal = 1
                '            Else
                '                row.Kredit = Me.GetSaldoHeaderCetakan(row.Primary, _
                '                row.Level_Perkiraan, TBL, row.Saldo_Normal)
                '                row.Tebal = 1
                '            End If
                '        End If
                '    Next

            End If

            Report.SetDataSource(DS)
            Report.SetParameterValue(Report.Parameter_FormatDay.ParameterFieldName, FormatDay)
            Report.SetParameterValue(Report.Parameter_FormatMonth.ParameterFieldName, FormatMonth)
            Report.SetParameterValue(Report.Parameter_FormatYear.ParameterFieldName, FormatYears)
            Report.SetParameterValue(Report.Parameter_FormatDecimal.ParameterFieldName, FormatDecimal)

            If ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
            Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets) = "" Then
                Negatif = 1
            Else
                Negatif = ObjSetting.GetSettingValue(Core_Module.ClsTransaksiSetting.EnumTipeTransaksi.Reporting, _
                Core_Module.ClsTransaksiSetting.EnumNamaParameter.UseBrackets)
            End If
            Report.SetParameterValue(Report.Parameter_Negatif.ParameterFieldName, Negatif)

            CRV.ReportSource = Report
            CRV.Show()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DS.Dispose()
            DTA.Dispose()
            ObjCur = Nothing
            ObjSetting = Nothing
        End Try
    End Function

    ' Private Function GetSaldoHeaderCetakan(ByVal vPrimaryAkun As Double, ByVal vLevel As Integer, _
    'ByVal TBL As DseMasterAkun.tblCetak_master_akunDataTable, ByVal vSaldoNormal As Integer) As Decimal
    '     Dim row As DseMasterAkun.tblCetak_master_akunRow
    '     Dim Filter As String = ""
    '     Dim DataView As DataRowView
    '     Dim retval As Decimal = 0

    '     Try
    '         If vLevel = 1 Then
    '             Filter = "[Primary Lvl 1] = " + vPrimaryAkun.ToString
    '         ElseIf vLevel = 2 Then
    '             Filter = "[Primary Lvl 2] = " + vPrimaryAkun.ToString
    '         Else
    '             Filter = "[Primary Lvl 3] = " + vPrimaryAkun.ToString
    '         End If

    '         TBL.DefaultView.RowFilter = Filter
    '         If TBL.DefaultView.Count > 0 Then
    '             For Each DataView In TBL.DefaultView
    '                 row = DataView.Row
    '                 If vSaldoNormal = 0 Then
    '                     retval += row.Debet
    '                 Else
    '                     retval += row.Kredit
    '                 End If
    '             Next
    '         End If

    '         TBL.DefaultView.RowFilter = ""
    '         Return retval
    '     Catch ex As Exception
    '         MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '     End Try
    ' End Function

    Public Function IsiDGVItem(ByVal vDGVItem As DataGridView, _
        ByVal Frm As FrmMasterAkun, ByVal vDefPrec As Integer) As Boolean
        Dim DTA As New DseMasterAkunTableAdapters.tblCariDetail_master_akunTableAdapter
        Dim TBL As New DseMasterAkun.tblCariDetail_master_akunDataTable
        Dim row As DseMasterAkun.tblCariDetail_master_akunRow
        Dim ObjCD As New Core_Module.ClsCekData
        Dim rowDGV As DataGridViewRow
        Dim DefaultPrecision As Integer = 0
        Dim i As Integer = 0
        Dim j As EnumTipeAkun2 = EnumTipeAkun2.Account_Payable

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            'kosongkan dgv
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            vDGVItem.Rows.Clear()

            'isikan tabel
            If DTA.Fill(TBL) > 0 Then
                'For Each row In TBL.Rows
                '    With row
                '        If .Header = 0 Then
                '            If .Saldo_Normal = 0 Then
                '                .Debet = Math.Round(.Beginning_Balance + .Debet - .Kredit, vDefPrec, MidpointRounding.AwayFromZero)
                '                .Kredit = 0
                '            Else
                '                .Kredit = Math.Round(.Beginning_Balance + _
                '                .Kredit - .Debet, vDefPrec, MidpointRounding.AwayFromZero)
                '                .Debet = 0
                '            End If
                '        End If
                '    End With
                'Next

                'For Each row In TBL.Rows
                '    If row.Header = 1 Then
                '        If row.Saldo_Normal = 0 Then
                '            row.Debet = Me.GetSaldoHeader(row.Primary, _
                '            row.Level_Perkiraan, TBL, row.Saldo_Normal)
                '        Else
                '            row.Kredit = Me.GetSaldoHeader(row.Primary, _
                '            row.Level_Perkiraan, TBL, row.Saldo_Normal)
                '        End If
                '    End If
                'Next

                For Each row In TBL.Rows
                    With row
                        vDGVItem.Rows.Add("")
                        rowDGV = vDGVItem.Rows(vDGVItem.Rows.Count - 1)
                        rowDGV.Cells(Frm.ColUrutan.Index).Value = .Urutan
                        rowDGV.Cells(Frm.ColPrimaryAkun.Index).Value = .Primary
                        'rowDGV.Cells(Frm.ColDebet.Index).Value = ObjCD.DigitGrouping(.Debet, vDefPrec)
                        'rowDGV.Cells(Frm.ColKredit.Index).Value = ObjCD.DigitGrouping(.Kredit, vDefPrec)
                        rowDGV.Cells(Frm.ColAccount.Index).Value = .Tampilan
                        rowDGV.Cells(Frm.ColPrimaryLvl1.Index).Value = .Primary_Lvl_1
                        rowDGV.Cells(Frm.ColPrimaryLvl2.Index).Value = .Primary_Lvl_2
                        rowDGV.Cells(Frm.ColPrimaryLvl3.Index).Value = .Primary_Lvl_3
                        rowDGV.Cells(Frm.ColLevel.Index).Value = .Level_Perkiraan
                        rowDGV.Cells(Frm.ColHeader.Index).Value = .Header
                        rowDGV.Cells(Frm.ColTipePerkiraan.Index).Value = .Tipe_Perkiraan
                        rowDGV.Cells(Frm.ColSaldoNormal.Index).Value = .Saldo_Normal
                        j = .Tipe_Perkiraan2
                        rowDGV.Cells(Frm.ColTipeAkun2Str.Index).Value = Me.GetTipeAkun2(j)
                        rowDGV.Cells(Frm.ColTipeAkun2.Index).Value = .Tipe_Perkiraan2
                        rowDGV.Cells(Frm.ColPrimaryLvl4.Index).Value = .Primary_Lvl_4
                        rowDGV.Cells(Frm.ColPrimaryLvl5.Index).Value = .Primary_Lvl_5
                        rowDGV.Cells(Frm.ColPrimaryLvl6.Index).Value = .Primary_Lvl_6
                        rowDGV.Cells(Frm.ColPrimaryLvl7.Index).Value = .Primary_Lvl_7
                    End With
                Next
            End If
            vDGVItem.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            vDGVItem.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            'nilai balikan 
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
            TBL = Nothing
            ObjCD = Nothing
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
        End Try
    End Function

    'Private Function GetSaldoHeader(ByVal vPrimaryAkun As Double, ByVal vLevel As Integer, _
    'ByVal TBL As DseMasterAkun.tblCariDetail_master_akunDataTable, ByVal vSaldoNormal As Integer) As Decimal
    '    Dim row As DseMasterAkun.tblCariDetail_master_akunRow
    '    Dim Filter As String = ""
    '    Dim DataView As DataRowView
    '    Dim retval As Decimal = 0

    '    Try
    '        If vLevel = 1 Then
    '            Filter = "[Primary Lvl 1] = " + vPrimaryAkun.ToString
    '        ElseIf vLevel = 2 Then
    '            Filter = "[Primary Lvl 2] = " + vPrimaryAkun.ToString
    '        Else
    '            Filter = "[Primary Lvl 3] = " + vPrimaryAkun.ToString
    '        End If

    '        TBL.DefaultView.RowFilter = Filter
    '        If TBL.DefaultView.Count > 0 Then
    '            For Each DataView In TBL.DefaultView
    '                row = DataView.Row
    '                If vSaldoNormal = 0 Then
    '                    retval += row.Debet
    '                Else
    '                    retval += row.Kredit
    '                End If
    '            Next
    '        End If

    '        TBL.DefaultView.RowFilter = ""
    '        Return retval
    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
    '    End Try
    'End Function

    Public Function GetDescription(ByVal vKodeAkun As String) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc_ByAccCode(vKodeAkun)) = False Then
                Return DTA.ScalarQuery_GetDesc_ByAccCode(vKodeAkun)
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

    Public Function GetDescription_ByPrimaryAcc(ByVal vPrimary_Acc As String) As String
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If String.IsNullOrEmpty(DTA.ScalarQuery_GetDesc(vPrimary_Acc)) = False Then
                Return DTA.ScalarQuery_GetDesc(vPrimary_Acc)
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

    Public Function GetJml(ByVal vKodeAkun As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml(vKodeAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJml(vKodeAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJml_NonHeader_Aktif(ByVal vKodeAkun As String) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetJml_Aktif_NonHeader(vKodeAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetJml_Aktif_NonHeader(vKodeAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
            Return ""
        Finally
            DTA.Dispose()
        End Try
    End Function

    ''' <summary>
    ''' mengembalikan apakah akun tertentu adalah akun header
    ''' jika akun tersebut tidak terdapat dalam database maka dianggap sebagai header
    ''' </summary>
    ''' <param name="vPrimaryAkun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetHeader(ByVal vPrimaryAkun As Double) As Integer
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            If DTA.ScalarQuery_GetHeader(vPrimaryAkun).HasValue = True Then
                Return DTA.ScalarQuery_GetHeader(vPrimaryAkun).Value
            Else
                Return 0
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        Finally
            DTA.Dispose()
        End Try
    End Function

    Public Function GetJmlAnak_Aktif(ByVal vPrimaryAkun As Double) As Double
        Dim DTA As New DseMasterAkunTableAdapters.tblmaster_akunTableAdapter
        Dim Lvl As Integer = -1

        Try
            DTA.Connection.ConnectionString = Core_Module.GetConstring
            'cek level
            If DTA.ScalarQuery_GetLevel(vPrimaryAkun).HasValue = True Then
                If Lvl = 1 Then
                    'jika level 1
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl1(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl1(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 2 Then
                    'jika level 2
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl2(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl2(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 3 Then
                    'jika level 3
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl3(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl3(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 4 Then
                    'jika level 4
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl4(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl4(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 5 Then
                    'jika level 5
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl5(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl5(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 6 Then
                    'jika level 6
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl6(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl6(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                ElseIf Lvl = 7 Then
                    'jika level 7
                    If DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl7(vPrimaryAkun).HasValue = True Then
                        Return DTA.ScalarQuery_GetJmlAktif_ByPrimaryLvl7(vPrimaryAkun).Value
                    Else
                        Return 0
                    End If
                Else
                    'jika level 8, ga mungkin punya anak
                    Return 0
                End If
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
