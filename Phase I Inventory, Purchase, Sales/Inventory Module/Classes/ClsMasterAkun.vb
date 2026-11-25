Option Explicit On
Option Strict Off

Public Class ClsMasterAkun
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
    '   Private mSaldoNormal As EnumSaldoNormal = EnumSaldoNormal.Debet
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
End Class
