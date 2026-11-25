Option Explicit On
Option Strict Off
Public Class ClsTransaksiCreditNote
#Region "Declaration"
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
    Dim objCD As New Core_Module.ClsCekData

    Public Enum EnumTipeTransaksi
        CREDIT_NOTE
        DEBIT_NOTE
    End Enum
    Public Enum EnumTipeKartu
        Customer
        Supplier
    End Enum
#End Region
End Class
